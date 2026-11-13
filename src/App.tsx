
import { useState } from 'react'
import Papa from 'papaparse'
import ExcelJS from 'exceljs'
import Encoding from 'encoding-japanese'
import './App.css'

type CSVFile = {
  id: number
  name: string
  encoding: string
  headers: string[]
  rows: Record<string, string>[]
  mappings: Record<string, string>
}

function App() {
  const [files, setFiles] = useState<CSVFile[]>([])
  const [error, setError] = useState('')
  const [outputFileName, setOutputFileName] = useState('結合結果')

  // 全ファイルの列名を統合
  const headers = Array.from(
   new Set(
     files.flatMap(file =>
       file.headers.map(header =>
         file.mappings[header]?.trim() || header
       )
     )
   )
 )


  // 全ファイルのデータを統合
  const rows = files.flatMap(file =>
   file.rows.map(row =>
     Object.fromEntries(
       file.headers.map(header => [
         file.mappings[header]?.trim() || header,
         row[header] ?? '',
       ])
     )
   )
 )

  // 文字コードを判定してCSVを読み込む
  const handleFileUpload = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const selectedFiles = Array.from(event.target.files ?? [])

    if (selectedFiles.length === 0) return

    setError('')

    try {
      const results = await Promise.all(
        selectedFiles.map(async file => {
          // ファイルをバイナリとして読み込む
          const buffer = await file.arrayBuffer()
          const bytes = new Uint8Array(buffer)

          // 文字コードを自動判定
          const detected = Encoding.detect(bytes)

          if (
            !detected ||
            detected === 'BINARY' ||
            detected === 'UTF32'
          ) {
            throw new Error(
              `${file.name}: 文字コードを判定できませんでした`
            )
          }

          // ASCIIはUTF-8として扱える
          const sourceEncoding =
            detected === 'ASCII' ? 'UTF8' : detected

          // Unicode文字列に変換
          const text = Encoding.convert(bytes, {
            to: 'UNICODE',
            from: sourceEncoding,
            type: 'string',
          }) as string

          // UTF-8 BOMなどを除去
          const cleanText = text.replace(/^\uFEFF/, '')

          // CSVを解析
          const result = Papa.parse<Record<string, string>>(
            cleanText,
            {
              header: true,
              skipEmptyLines: true,
              dynamicTyping: false,
            }
          )

          if (result.errors.length > 0) {
            throw new Error(
              `${file.name}: ${result.errors[0].message}`
            )
          }

          const parsedHeaders = result.meta.fields ?? []

          return {
            id: Date.now() + Math.random(),
            name: file.name,
            encoding: sourceEncoding,
            headers: parsedHeaders,
            rows: result.data,
            mappings: Object.fromEntries(
              parsedHeaders.map(header => [header, header])
            ),
          }
        })
      )

      setFiles(previous => [...previous, ...results])

    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : '読み込みに失敗しました'
      )
    }

    event.target.value = ''
  }

  // 個別ファイルを削除
  const deleteFile = (id: number) => {
    setFiles(previous =>
      previous.filter(file => file.id !== id)
    )
  }

  const updateMapping = (
  fileId: number,
  originalHeader: string,
  newHeader: string
) => {
  setFiles(previous =>
    previous.map(file =>
      file.id === fileId
        ? {
            ...file,
            mappings: {
              ...file.mappings,
              [originalHeader]: newHeader,
            },
          }
        : file
    )
  )
}

  // 全ファイルを削除
  const clearFiles = () => {
    setFiles([])
  }

  // 結合CSVをダウンロード
  const downloadCSV = () => {
    if (rows.length === 0) return

    const csv = Papa.unparse({
      fields: headers,
      data: rows.map(row =>
        headers.map(header => row[header] ?? '')
      ),
    })

    // UTF-8 BOMを付ける
    const blob = new Blob(['\uFEFF' + csv], {
      type: 'text/csv;charset=utf-8;',
    })

    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')

    // ファイル名を整える
    const baseName = outputFileName
      .trim()
      .replace(/\.csv$/i, '')
      .replace(/[\\/:*?"<>|]/g, '_')

    link.href = url
    link.download = `${baseName || '結合結果'}.csv`
    link.click()

    URL.revokeObjectURL(url)
  }

 // Excel形式でダウンロード
  const downloadExcel = async () => {
    if (rows.length === 0) return

    try {
      const workbook = new ExcelJS.Workbook()
      const worksheet = workbook.addWorksheet('結合データ')

      // 列幅を設定
      worksheet.columns = headers.map(header => ({
        width: Math.min(Math.max(header.length + 5, 12), 30),
      }))

      // ヘッダー行
      const headerRow = worksheet.addRow(headers)

      headerRow.font = { bold: true }
      headerRow.eachCell(cell => {
        cell.numFmt = '@'
      })

      // データ行を追加
      rows.forEach(row => {
        const newRow = worksheet.addRow(
          headers.map(header => String(row[header] ?? ''))
        )

        // すべてのセルを文字列として扱う
        newRow.eachCell(cell => {
          cell.value = String(cell.value ?? '')
          cell.numFmt = '@'
        })
      })

      // Excelファイルを生成
      const buffer = await workbook.xlsx.writeBuffer()

      const blob = new Blob([buffer as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      })

      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')

      // ファイル名を整える
      const baseName = outputFileName
        .trim()
        .replace(/\.(csv|xlsx)$/i, '')
        .replace(/[\\/:*?"<>|]/g, '_')

      link.href = url
      link.download = `${baseName || '結合結果'}.xlsx`

      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)

      // ダウンロード開始後にURLを解放
      window.setTimeout(() => {
        URL.revokeObjectURL(url)
      }, 1000)


    } catch (err) {
      console.error('Excel出力エラー:', err)

      setError(
        err instanceof Error
          ? `Excel出力エラー：${err.message}`
          : 'Excelファイルの出力に失敗しました'
      )
    }
  }

  return (
    <main className="app">
      <header className="header">
        <h1>CSV Workbench</h1>
        <p>CSVデータを簡単に読み込み・集計・整理</p>
      </header>

      <section className="upload-section">
        <h2>CSVファイルを読み込む</h2>

        <label className="upload-button">
          CSVファイルを選択
          <input
            type="file"
            accept=".csv,text/csv"
            multiple
            onChange={handleFileUpload}
          />
        </label>

        {error && (
          <p className="error">エラー：{error}</p>
        )}
      </section>

      {files.length > 0 && (
        <section className="file-section">
          <div className="data-heading">
            <div>
              <h2>読み込んだファイル</h2>
              <span>{files.length} ファイル</span>
            </div>

            <button
              className="clear-button"
              onClick={clearFiles}
            >
              すべて削除
            </button>
          </div>

          <ul className="file-list">
            {files.map(file => (
              <li key={file.id}>
                <span className="file-icon">CSV</span>

                <div className="file-info">
                  <span className="file-name">
                    {file.name}
                  </span>
                  <span className="file-count">
                    {file.rows.length} 行
                    {' / '}
                    {file.encoding}
                  </span>
                </div>

                <button
                  className="delete-button"
                  onClick={() => deleteFile(file.id)}
                >
                  削除
                </button>
              </li>
            ))}
          </ul>
        </section>
      )}

       <section className="mapping-section">
         <h2>列名の対応付け</h2>
         <p>同じ列として扱いたい項目は、出力時の列名を統一してください。</p>

         {files.map(file => (
           <div className="mapping-file" key={file.id}>
             <h3>{file.name}</h3>

             {file.headers.map(header => (
               <div className="mapping-row" key={header}>
                 <label>{header}</label>
                 <span>→</span>
                 <input
                   type="text"
                   value={file.mappings[header] ?? header}
                   onChange={event =>
                     updateMapping(file.id, header, event.target.value)
                   }
                 />
               </div>
              ))}
           </div>
         ))}
       </section>

      {files.length > 0 && (
        <section className="data-section">
          <div className="data-heading">
            <div>
              <h2>結合データ一覧</h2>
              <span>
                {rows.length} 行 / {headers.length} 列
              </span>
            </div>
          </div>

          {rows.length > 0 ? (
            <>
              <div className="filename-section">
                <label htmlFor="output-filename">
                  ダウンロードファイル名
                </label>

                <div className="filename-input-wrapper">
                  <input
                    id="output-filename"
                    type="text"
                    value={outputFileName}
                    onChange={event =>
                      setOutputFileName(event.target.value)
                    }
                    placeholder="結合結果"
                  />
                  <span>.csv</span>
                </div>
              </div>

              <div className="download-area">
                <button
                   className="download-button"
                   onClick={downloadCSV}
                >
                  CSVをダウンロード
                </button>

                <button
                   className="download-button excel-download-button"
                   onClick={downloadExcel}
                >
                  Excel（.xlsx）をダウンロード
                </button>
              </div>

              <div className="table-container">
                <table>
                  <thead>
                    <tr>
                      {headers.map(header => (
                        <th key={header}>{header}</th>
                      ))}
                    </tr>
                  </thead>

                  <tbody>
                    {rows.slice(0, 100).map((row, index) => (
                      <tr key={index}>
                        {headers.map(header => (
                          <td key={header}>
                            {row[header] ?? ''}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {rows.length > 100 && (
                <p className="note">
                  最初の100行を表示しています。
                  全データはダウンロードできます。
                </p>
              )}
            </>
          ) : (
            <p className="note">
              データがありません。
            </p>
          )}
        </section>
      )}
    </main>
  )
}

export default App
