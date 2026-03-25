import React, { useEffect, useMemo, useRef, useState } from "react";
import { HotTable } from "@handsontable/react";
import { registerAllModules } from "handsontable/registry";
import "handsontable/dist/handsontable.full.min.css";

registerAllModules();

const API_BASE_URL =
  (import.meta.env.VITE_API_BASE_URL || "http://localhost:5000").replace(/\/$/, "");

const DEFAULT_ROWS = 18;
const DEFAULT_COLS = 6;
const DEFAULT_OCR_ATTEMPTS = 1;
const EXPORT_MODE_SEPARATE = "separate_sheets";
const EXPORT_MODE_COMBINED = "combined_sheet";

function columnLabel(index) {
  let label = "";
  let current = index;

  do {
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26) - 1;
  } while (current >= 0);

  return label;
}

function ensureGridShape(grid, minRows = DEFAULT_ROWS, minCols = DEFAULT_COLS) {
  const currentRows = Math.max(grid.length, minRows);
  const currentCols = Math.max(
    minCols,
    grid.reduce((max, row) => Math.max(max, row.length), 0)
  );

  return Array.from({ length: currentRows }, (_, rowIndex) =>
    Array.from({ length: currentCols }, (_, colIndex) => grid[rowIndex]?.[colIndex] ?? "")
  );
}

function cellsToGrid(cells, meta = {}) {
  if (!cells?.length) {
    return ensureGridShape([], DEFAULT_ROWS, DEFAULT_COLS);
  }

  const maxRow = Math.max(...cells.map((cell) => cell.row), 0);
  const maxCol = Math.max(...cells.map((cell) => cell.col), 0);
  const grid = Array.from({ length: maxRow + 1 }, () =>
    Array.from({ length: maxCol + 1 }, () => "")
  );

  cells.forEach((cell) => {
    if (!grid[cell.row]) {
      grid[cell.row] = Array.from({ length: maxCol + 1 }, () => "");
    }
    grid[cell.row][cell.col] = cell.text ?? "";
  });

  return ensureGridShape(
    grid,
    Math.max(DEFAULT_ROWS, meta.recognized_rows || 0),
    Math.max(DEFAULT_COLS, meta.recognized_columns || 0)
  );
}

function cloneGrid(grid) {
  return grid.map((row) => [...row]);
}

function countFilledCells(grid) {
  return grid.reduce(
    (total, row) => total + row.filter((cell) => String(cell ?? "").trim() !== "").length,
    0
  );
}

function App() {
  const videoRef = useRef(null);
  const canvasRef = useRef(null);
  const fileInputRef = useRef(null);
  const tableRefs = useRef({});
  const streamRef = useRef(null);

  const [receipts, setReceipts] = useState([]);
  const [activeReceiptId, setActiveReceiptId] = useState(null);
  const [loading, setLoading] = useState(false);
  const [cameraActive, setCameraActive] = useState(false);
  const [ocrAttempts, setOcrAttempts] = useState(DEFAULT_OCR_ATTEMPTS);
  const [exportMode, setExportMode] = useState(EXPORT_MODE_COMBINED);
  const [statusMessage, setStatusMessage] = useState("準備完了");
  const [errorMessage, setErrorMessage] = useState("");

  useEffect(() => {
    return () => {
      if (streamRef.current) {
        streamRef.current.getTracks().forEach((track) => track.stop());
      }
    };
  }, []);

  const activeReceipt = useMemo(
    () => receipts.find((receipt) => receipt.id === activeReceiptId) ?? null,
    [receipts, activeReceiptId]
  );

  const syncReceiptGrid = (receiptId, nextGrid) => {
    setReceipts((previous) =>
      previous.map((receipt) =>
        receipt.id === receiptId
          ? { ...receipt, grid: ensureGridShape(cloneGrid(nextGrid)) }
          : receipt
      )
    );
  };

  const syncFromTable = (receiptId) => {
    const tableRef = tableRefs.current[receiptId];
    const instance = tableRef?.hotInstance;
    if (!instance) {
      return;
    }

    syncReceiptGrid(receiptId, instance.getData());
  };

  const startCamera = async () => {
    if (!navigator.mediaDevices?.getUserMedia) {
      setErrorMessage("このブラウザではカメラ撮影に対応していません。");
      return;
    }

    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: {
          facingMode: { ideal: "environment" },
        },
        audio: false,
      });

      streamRef.current = stream;

      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await videoRef.current.play();
      }

      setCameraActive(true);
      setErrorMessage("");
      setStatusMessage("カメラ起動");
    } catch (error) {
      console.error(error);
      setErrorMessage("カメラにアクセスできませんでした。権限を許可するか、画像アップロードを使ってください。");
    }
  };

  const stopCamera = () => {
    if (streamRef.current) {
      streamRef.current.getTracks().forEach((track) => track.stop());
      streamRef.current = null;
    }

    if (videoRef.current) {
      videoRef.current.srcObject = null;
    }

    setCameraActive(false);
    setStatusMessage("カメラ停止");
  };

  const addReceipt = ({ grid, previewUrl, source, meta }) => {
    setReceipts((previous) => {
      const nextIndex = previous.length + 1;
      const receipt = {
        id: Date.now() + nextIndex,
        name: `Receipt ${nextIndex}`,
        grid,
        previewUrl: previewUrl || null,
        source,
        meta: meta || {},
        createdAt: new Date().toISOString(),
      };
      const nextReceipts = [...previous, receipt];
      setActiveReceiptId(receipt.id);
      return nextReceipts;
    });
  };

  const processReceiptBlob = async (blob, source, previewUrl) => {
    setLoading(true);
    setErrorMessage("");
    setStatusMessage("OCR実行中");

    try {
      const formData = new FormData();
      formData.append("image", blob, "receipt.png");
      formData.append("ocr_attempts", String(ocrAttempts));

      const response = await fetch(`${API_BASE_URL}/api/ocr`, {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error(`OCR request failed with status ${response.status}`);
      }

      const payload = await response.json();
      const grid = cellsToGrid(payload.cells || [], payload.meta || {});

      addReceipt({
        grid,
        previewUrl,
        source,
        meta: payload.meta || {},
      });

      setStatusMessage(
        payload.cells?.length
          ? `OCR完了: ${payload.cells.length}セル検出`
          : "OCR完了: 文字検出なし"
      );
    } catch (error) {
      console.error(error);
      setErrorMessage("OCRに失敗しました。明るい画像で再試行するか、空シートで手入力してください。");
      setStatusMessage("OCR失敗");
      if (previewUrl) {
        URL.revokeObjectURL(previewUrl);
      }
    } finally {
      setLoading(false);
    }
  };

  const captureReceipt = async () => {
    if (!videoRef.current || !canvasRef.current || !cameraActive) {
      setErrorMessage("撮影前のカメラ開始");
      return;
    }

    const video = videoRef.current;
    const canvas = canvasRef.current;

    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;

    const context = canvas.getContext("2d");
    context.drawImage(video, 0, 0, canvas.width, canvas.height);

    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png", 1));
    if (!blob) {
      setErrorMessage("カメラ画像生成不可");
      return;
    }

    const previewUrl = URL.createObjectURL(blob);
    await processReceiptBlob(blob, "camera", previewUrl);
  };

  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    const previewUrl = URL.createObjectURL(file);
    await processReceiptBlob(file, "upload", previewUrl);
    event.target.value = "";
  };

  const removeReceipt = (receiptId) => {
    setReceipts((previous) => {
      const target = previous.find((receipt) => receipt.id === receiptId);
      if (target?.previewUrl) {
        URL.revokeObjectURL(target.previewUrl);
      }

      const nextReceipts = previous.filter((receipt) => receipt.id !== receiptId);

      if (receiptId === activeReceiptId) {
        setActiveReceiptId(nextReceipts[0]?.id ?? null);
      }

      return nextReceipts;
    });
  };

  const duplicateReceipt = () => {
    if (!activeReceipt) {
      return;
    }

    setReceipts((previous) => {
      const clone = {
        ...activeReceipt,
        id: Date.now(),
        name: `${activeReceipt.name} Copy`,
        grid: cloneGrid(activeReceipt.grid),
        previewUrl: activeReceipt.previewUrl,
        createdAt: new Date().toISOString(),
      };
      const nextReceipts = [...previous, clone];
      setActiveReceiptId(clone.id);
      return nextReceipts;
    });
    setStatusMessage("シート複製");
  };

  const renameReceipt = (receiptId, value) => {
    setReceipts((previous) =>
      previous.map((receipt) =>
        receipt.id === receiptId ? { ...receipt, name: value || "Untitled Receipt" } : receipt
      )
    );
  };

  const downloadExcel = async () => {
    if (!receipts.length) {
      return;
    }

    setLoading(true);
    setErrorMessage("");
    setStatusMessage("Excel作成中");

    try {
      const rows = receipts.flatMap((receipt, receiptIndex) =>
        receipt.grid.flatMap((row, rowIndex) =>
          row.map((text, colIndex) => ({
            receipt_no: receiptIndex + 1,
            receipt_name: receipt.name,
            row: rowIndex,
            col: colIndex,
            text,
          }))
        )
      );

      const response = await fetch(`${API_BASE_URL}/api/export_excel`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ rows, export_mode: exportMode }),
      });

      if (!response.ok) {
        throw new Error(`Export failed with status ${response.status}`);
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      const today = new Date();
      const yyyy = String(today.getFullYear());
      const mm = String(today.getMonth() + 1).padStart(2, "0");
      const dd = String(today.getDate()).padStart(2, "0");
      link.href = url;
      link.download = `${yyyy}${mm}${dd}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
      setStatusMessage("Excel出力");
    } catch (error) {
      console.error(error);
      setErrorMessage("Excelの出力に失敗しました。");
      setStatusMessage("出力失敗");
    } finally {
      setLoading(false);
    }
  };

  const stats = activeReceipt
    ? {
        rows: activeReceipt.grid.length,
        cols: activeReceipt.grid[0]?.length || 0,
        filled: countFilledCells(activeReceipt.grid),
      }
    : null;

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-title">
          <h1>レシートからExcelへ</h1>
        </div>
        <div className="hero-actions">
          <select
            value={exportMode}
            onChange={(event) => setExportMode(event.target.value)}
            disabled={loading}
            className="hero-select"
          >
            <option value={EXPORT_MODE_SEPARATE}>レシートごとに別シート出力</option>
            <option value={EXPORT_MODE_COMBINED}>全レシートを1シートに集約</option>
          </select>
          <button className="primary-button" onClick={downloadExcel} disabled={!receipts.length || loading}>
            Excelを書き出す
          </button>
        </div>
      </header>

      <div className="workspace">
        <aside className="capture-panel">
          <section className="panel-card">
            <div className="panel-heading">
              <h2>取り込み</h2>
              <span className="status-chip">{loading ? "処理中" : "待機中"}</span>
            </div>

            <div className="camera-frame">
              <video ref={videoRef} muted playsInline />
              {!cameraActive && <div className="camera-placeholder">カメラプレビュー</div>}
            </div>

            <canvas ref={canvasRef} className="hidden-canvas" />
            <input
              ref={fileInputRef}
              type="file"
              accept="image/*"
              className="hidden-input"
              onChange={handleFileChange}
            />

            <div className="message-stack">
              <p className="message-label">OCR試行回数</p>
              <select
                value={ocrAttempts}
                onChange={(event) => setOcrAttempts(Number(event.target.value))}
                disabled={loading}
                className="ocr-attempt-select"
              >
                <option value={1}>1回</option>
                <option value={2}>2回</option>
                <option value={3}>3回</option>
                <option value={4}>4回</option>
                <option value={5}>5回</option>
              </select>
            </div>

            <div className="button-stack">
              <div className="button-row balanced">
                <button className="primary-button" onClick={startCamera} disabled={cameraActive || loading}>
                  カメラ開始
                </button>
                <button className="secondary-button" onClick={stopCamera} disabled={!cameraActive || loading}>
                  カメラ停止
                </button>
              </div>

              <div className="button-row balanced">
                <button className="accent-button" onClick={captureReceipt} disabled={!cameraActive || loading}>
                  撮影・OCR
                </button>
                <button className="secondary-button" onClick={handleUploadClick} disabled={loading}>
                  画像選択
                </button>
              </div>
            </div>

            <div className="message-stack">
              <p className="message-label">状態</p>
              <p className="message-text">{statusMessage}</p>
              {errorMessage && <p className="error-text">{errorMessage}</p>}
            </div>
          </section>

          <section className="panel-card">
            <div className="panel-heading">
              <h2>操作のヒント</h2>
            </div>
            <ul className="tips-list">
              <li>キーボードでのコピー&ペースト</li>
              <li>右クリックでの挿入・削除・元に戻す操作</li>
              <li>行見出しと列見出しのドラッグ並べ替え</li>
              <li>出力方式の選択</li>
            </ul>
          </section>
        </aside>

        <main className="editor-panel">
          <section className="receipt-browser">
            <div className="panel-heading">
              <h2>レシート一覧</h2>
              <span className="receipt-count">{receipts.length}</span>
            </div>

            {receipts.length === 0 && (
              <div className="empty-state">
                <p>まだレシートなし。</p>
                <p>カメラ撮影、画像アップロードからの開始。</p>
              </div>
            )}

            <div className="receipt-list">
              {receipts.map((receipt) => {
                const filledCells = countFilledCells(receipt.grid);
                const isActive = receipt.id === activeReceiptId;

                return (
                  <button
                    type="button"
                    key={receipt.id}
                    className={`receipt-card ${isActive ? "active" : ""}`}
                    onClick={() => setActiveReceiptId(receipt.id)}
                  >
                    <div className="receipt-card-header">
                      <strong>{receipt.name}</strong>
                      <span>{filledCells} cells</span>
                    </div>
                    <div className="receipt-card-meta">
                      <span>{receipt.source}</span>
                      <span>{receipt.meta.variant || "manual"}</span>
                    </div>
                    {receipt.previewUrl ? (
                      <img className="receipt-thumb" src={receipt.previewUrl} alt={receipt.name} />
                    ) : (
                      <div className="receipt-thumb placeholder">画像なし</div>
                    )}
                  </button>
                );
              })}
            </div>
          </section>

          <section className="sheet-panel">
            {activeReceipt ? (
              <>
                <div className="sheet-toolbar">
                  <div className="sheet-title-group">
                    <input
                      className="sheet-title-input"
                      value={activeReceipt.name}
                      onChange={(event) => renameReceipt(activeReceipt.id, event.target.value)}
                    />
                    <div className="sheet-stats">
                      <span>{stats?.rows} rows</span>
                      <span>{stats?.cols} columns</span>
                      <span>{stats?.filled} filled cells</span>
                    </div>
                  </div>

                  <div className="button-row compact">
                    <button className="secondary-button" onClick={duplicateReceipt} disabled={loading}>
                      複製
                    </button>
                    <button className="danger-button" onClick={() => removeReceipt(activeReceipt.id)} disabled={loading}>
                      削除
                    </button>
                  </div>
                </div>

                <div className="sheet-summary">
                  <span>OCRバリエーション: {activeReceipt.meta.variant || "manual"}</span>
                  <span>推定行数: {activeReceipt.meta.recognized_rows || activeReceipt.grid.length}</span>
                  <span>検出トークン数: {activeReceipt.meta.recognized_tokens || countFilledCells(activeReceipt.grid)}</span>
                  <span>切り出し: {activeReceipt.meta.cropped ? "あり" : "なし"}</span>
                  <span>傾き補正: {activeReceipt.meta.deskew_angle ?? 0}°</span>
                  <span>OCR試行回数: {activeReceipt.meta.variant_runs ?? 1}</span>
                  <span>指定回数: {activeReceipt.meta.requested_attempts ?? 1}</span>
                </div>

                <div className="table-shell">
                  <HotTable
                    ref={(instance) => {
                      tableRefs.current[activeReceipt.id] = instance;
                    }}
                    data={activeReceipt.grid}
                    rowHeaders={true}
                    colHeaders={(index) => columnLabel(index)}
                    width="100%"
                    height={620}
                    licenseKey="non-commercial-and-evaluation"
                    stretchH="all"
                    manualRowMove={true}
                    manualColumnMove={true}
                    manualRowResize={true}
                    manualColumnResize={true}
                    minSpareRows={4}
                    minSpareCols={1}
                    contextMenu={true}
                    dropdownMenu={true}
                    filters={true}
                    columnSorting={true}
                    copyPaste={true}
                    search={true}
                    multiColumnSorting={true}
                    outsideClickDeselects={false}
                    autoWrapRow={true}
                    autoWrapCol={true}
                    afterChange={(changes, source) => {
                      if (!changes || source === "loadData") {
                        return;
                      }
                      syncFromTable(activeReceipt.id);
                    }}
                    afterCreateRow={() => syncFromTable(activeReceipt.id)}
                    afterRemoveRow={() => syncFromTable(activeReceipt.id)}
                    afterCreateCol={() => syncFromTable(activeReceipt.id)}
                    afterRemoveCol={() => syncFromTable(activeReceipt.id)}
                    afterRowMove={() => syncFromTable(activeReceipt.id)}
                    afterColumnMove={() => syncFromTable(activeReceipt.id)}
                  />
                </div>
              </>
            ) : (
              <div className="empty-state large">
                <p>編集対象レシートの選択。</p>
                <p>コピー&ペースト、行列並べ替え、右クリック編集への対応。</p>
              </div>
            )}
          </section>
        </main>
      </div>
    </div>
  );
}

export default App;
