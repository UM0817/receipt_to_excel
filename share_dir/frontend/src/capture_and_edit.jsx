import React, { useRef, useState } from "react";

function buildGrid(cells) {
  const maxRow = Math.max(...cells.map(c => c.row));
  const maxCol = Math.max(...cells.map(c => c.col));

  const grid = Array.from({ length: maxRow + 1 }, () =>
    Array.from({ length: maxCol + 1 }, () => "")
  );

  cells.forEach(c => {
    grid[c.row][c.col] = c.text;
  });

  return grid;
}

function App(){
  // HTMLの各要素にアクセスするための参照作成
  const videoRef = useRef(null);    // <video>要素との紐づけ
  const canvasRef = useRef(null);   // <canvas>要素との紐づけ

  // [状態,更新する関数]の定義
  const [receipts, setReceipts] = useState([]);   // 配列、初期値は空配列
  const [loading, setLoading] = useState(false);  // 通信中などの処理状態、初期値はfalse  

  // カメラ開始
  const startCamera = async () => {     // カメラを起動する非同期関数の定義
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: true });    // ブラウザのAPIで映像を取得
      videoRef.current.srcObject = stream;    // 取得した画像をvideoRefで紐付けた<video>データソースとしてセット
      videoRef.current.play();    // 画面に表示
    } catch (err) {
      alert("カメラの起動に失敗しました。権限を確認してください。");
      console.error(err);
    }
  };

  // 🔵 撮影ボタン
  const captureReceipt = async () => {
    setLoading(true);

    const video = videoRef.current;
    const canvas = canvasRef.current;

    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(video, 0, 0);

    canvas.toBlob(async (blob) => {
      const form = new FormData();
      form.append("image", blob, "receipt.png");

      const res = await fetch("http://localhost:5000/api/ocr", {
        method: "POST",
        body: form
      });

      const json = await res.json();

      if(json.cells){
        const newReceipt = {
          id: Date.now(),
          cells: json.cells
        };

        setReceipts(prev => [...prev, newReceipt]);
      }

      setLoading(false);
    }, "image/png");
  };

  // 編集
  const updateCell = (receiptId, row, col, value) => {

    setReceipts(prev =>
      prev.map(r =>
        r.id === receiptId
          ? {
              ...r,
              cells: r.cells.map(c =>
                c.row === row && c.col === col
                  ? { ...c, text: value }
                  : c
              )
            }
          : r
      )
    );

  };

  // 🔵 Excel出力（全部まとめて送る）
  const downloadExcel = async () => {

    const allData = receipts.flatMap((r, idx) =>
      r.cells.map(cell => ({
        receipt_no: idx + 1,
        row: cell.row,
        col: cell.col,
        text: cell.text
      }))
    );

    const res = await fetch("http://localhost:5000/api/export_excel", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({ rows: allData })
    });

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "receipts.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div style={{padding:20}}>
      <h2>レシート読み取り</h2>

      <video ref={videoRef} width={400}></video><br/>
      <button onClick={startCamera}>カメラ開始</button>
      <button onClick={captureReceipt} disabled={loading}>
        📸 撮影
      </button>

      <canvas ref={canvasRef} style={{display:'none'}}/>

      <hr/>

      {receipts.map((receipt, rIndex) => {

        const grid = buildGrid(receipt.cells);

        return (
          <div key={receipt.id} style={{marginBottom:40}}>
            <h3>レシート {rIndex+1}</h3>

            <table border="1" style={{borderCollapse:"collapse"}}>
              <tbody>

              {grid.map((row,rowIndex)=>(
                <tr key={rowIndex}>

                {row.map((cell,colIndex)=>{

                  const cellData = receipt.cells.find(
                    c => c.row===rowIndex && c.col===colIndex
                  );

                  return(
                    <td key={colIndex}>
                      <input
                        style={{width:120}}
                        value={cell}
                        onChange={(e)=>updateCell(
                          receipt.id,
                          rowIndex,
                          colIndex,
                          e.target.value
                        )}
                      />
                    </td>
                  )
                })}

                </tr>
              ))}

              </tbody>
            </table>

          </div>
        )

      })}

      {receipts.length > 0 && (
        <button onClick={downloadExcel}>
          Excelまとめてダウンロード
        </button>
      )}
    </div>
  );
}

export default App;