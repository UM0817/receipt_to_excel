import React, { useRef, useState } from "react";

function App(){
  const videoRef = useRef(null);
  const canvasRef = useRef(null);

  const [receipts, setReceipts] = useState([]);
  const [loading, setLoading] = useState(false);

  // カメラ開始
  const startCamera = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: true });
      videoRef.current.srcObject = stream;
      videoRef.current.play();
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

      if(json.lines){
        const newReceipt = {
          id: Date.now(),
          lines: json.lines
        };

        // 🔵 ここが追記処理
        setReceipts(prev => [...prev, newReceipt]);
      }

      setLoading(false);
    }, "image/png");
  };

  // 編集
  const updateLine = (receiptId, lineIndex, value) => {
    setReceipts(prev =>
      prev.map(r =>
        r.id === receiptId
          ? { ...r, lines: r.lines.map((l,i)=> i===lineIndex ? value : l) }
          : r
      )
    );
  };

  // 🔵 Excel出力（全部まとめて送る）
  const downloadExcel = async () => {

    const allData = receipts.flatMap((r, idx) =>
      r.lines.map(line => ({
        receipt_no: idx + 1,
        text: line
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

      {receipts.map((receipt, rIndex) => (
        <div key={receipt.id} style={{marginBottom:30}}>
          <h3>レシート {rIndex + 1}</h3>
          {receipt.lines.map((line, i) => (
            <div key={i}>
              <input
                style={{width:"80%"}}
                value={line}
                onChange={(e)=>updateLine(receipt.id, i, e.target.value)}
              />
            </div>
          ))}
        </div>
      ))}

      {receipts.length > 0 && (
        <button onClick={downloadExcel}>
          Excelまとめてダウンロード
        </button>
      )}
    </div>
  );
}

export default App;