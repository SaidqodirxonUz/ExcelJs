/* eslint-disable no-unused-vars */
import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

const tableStyle = {
  width: "100%",
  borderCollapse: "collapse",
  margin: "20px 0",
};

const cellStyle = {
  border: "1px solid #ddd",
  padding: "8px",
};

const headerCellStyle = {
  ...cellStyle,
  backgroundColor: "#00f",
  fontWeight: "bold",
};

const inputStyle = {
  border: "none",
  padding: "8px",
  width: "100%",
  boxSizing: "border-box",
};

const buttonStyle = {
  padding: "10px 20px",
  fontSize: "16px",
  marginTop: "20px",
  cursor: "pointer",
  backgroundColor: "#4CAF50",
  color: "white",
  border: "none",
  borderRadius: "4px",
};

function ExcelImportExport() {
  const [editingData, setEditingData] = useState([]);
  const [editedFileName, setEditedFileName] = useState("");

  const calculateCategoryTotals = (category, editingData) => {
    const categoryData = editingData.filter(
      (row) => row["Вид запасов"] === category
    );
    const categoryTotal = categoryData.reduce(
      (total, row) => {
        Object.keys(total).forEach((key) => {
          if (typeof row[key] === "number") {
            total[key] += row[key];
          }
        });
        return total;
      },
      { Number: "", Name: "", Type: category } // Starting with an object for the category totals
    );

    console.log(categoryTotal);
    return categoryTotal;
  };

  const handleEdit = (rowIndex, columnName, value) => {
    setEditingData((prevEditingData) => {
      const updatedData = [...prevEditingData];
      updatedData[rowIndex][columnName] = value;
      return updatedData;
    });
  };

  const handleSave = async () => {
    const uniqueCategories = [
      ...new Set(editingData.map((row) => row["Вид запасов"])),
    ];

    const updatedData = editingData.map((row, rowIndex) => {
      const category = row["Вид запасов"];
      const isLastRow = rowIndex === editingData.length - 1;

      if (isLastRow && uniqueCategories.includes(category)) {
        const categoryTotal = calculateCategoryTotals(category, editingData);

        // Qo'shilgan qator
        const newRow = {
          "№ п/п": "",
          Наименование: `Yig'indi: ${category}`,
          "Вид запасов": category,
          "Ед. изм": "",
          Количество: categoryTotal["Количество"],
          Цена: categoryTotal["Цена"],
          Сумма: categoryTotal["Сумма"],
          "Дата приобретения": "",
          Поставщик: "",
          ИНН: "",
          Примечание: "",
          Предложение: "",
        };

        // Har bir ustunga qo'shilgan qatordan avvalgi qator
        return [row, newRow].reduce(
          (acc, item) => {
            Object.keys(acc).forEach((key) => {
              if (typeof item[key] === "number") {
                acc[key] += item[key];
              }
            });
            return acc;
          },
          { Number: "", Name: "", Type: category }
        );
      }

      return row;
    });

    const editedWorkbook = new ExcelJS.Workbook();
    const editedSheet = editedWorkbook.addWorksheet("Sheet1");

    // Table headers
    editedSheet.columns = [
      { header: "№ п/п", key: "Number", style: headerCellStyle },
      { header: "Наименование", key: "Name", style: headerCellStyle },
      { header: "Вид запасов", key: "Type", style: headerCellStyle },
      { header: "Ед. изм", key: "Unit", style: headerCellStyle },
      { header: "Количество", key: "Quantity", style: headerCellStyle },
      { header: "Цена", key: "Price", style: headerCellStyle },
      { header: "Сумма", key: "Amount", style: headerCellStyle },
      {
        header: "Дата приобретения",
        key: "PurchaseDate",
        style: headerCellStyle,
      },
      { header: "Поставщик", key: "Supplier", style: headerCellStyle },
      { header: "ИНН", key: "INN", style: headerCellStyle },
      { header: "Примечание", key: "Note", style: headerCellStyle },
      { header: "Предложение", key: "Offer", style: headerCellStyle },
    ];

    editingData.forEach((row) => {
      editedSheet.addRow(row);
    });

    uniqueCategories.forEach((category) => {
      const totalRow = calculateCategoryTotals(category, editingData);
      editedSheet.addRow(totalRow);
    });

    const editedFileData = await editedWorkbook.xlsx.writeBuffer();

    const blob = new Blob([editedFileData], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = editedFileName || "edited_data.xlsx";
    a.click();
  };

  useEffect(() => {
    // Read the initial data from the "data.xlsx" file
    const fetchData = async () => {
      try {
        const response = await fetch("/src/arrangedData.xlsx");
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });

        const sheetName = workbook.SheetNames[0];
        const excelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        setEditingData([...excelData]);
      } catch (error) {
        console.error("Error fetching data:", error);
      }
    };

    fetchData();
  }, []);

  return (
    <div>
      <h1>Edit Excel Data</h1>

      <br />
      <button style={buttonStyle} onClick={handleSave}>
        Save to Excel
      </button>
      {editingData.length > 0 && (
        <table style={tableStyle}>
          <thead>
            <tr>
              <th style={headerCellStyle}>№ п/п</th>
              <th style={headerCellStyle}>Наименование</th>
              <th style={headerCellStyle}>Вид запасов</th>
              <th style={headerCellStyle}>Ед. изм</th>
              <th style={headerCellStyle}>Количество</th>
              <th style={headerCellStyle}>Цена</th>
              <th style={headerCellStyle}>Сумма</th>
              <th style={headerCellStyle}>Дата приобретения</th>
              <th style={headerCellStyle}>Поставщик</th>
              <th style={headerCellStyle}>ИНН</th>
              <th style={headerCellStyle}>Примечание</th>
              <th style={headerCellStyle}>Предложение</th>
            </tr>
          </thead>
          <tbody>
            {editingData.map((row, rowIndex) => (
              <React.Fragment key={rowIndex}>
                <tr>
                  <td style={cellStyle}>{row["№ п/п"]}</td>
                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Наименование"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "Наименование", e.target.value)
                      }
                    />
                  </td>
                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Вид запасов"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "Вид запасов", e.target.value)
                      }
                    />
                  </td>

                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Ед. изм"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "Ед. изм", e.target.value)
                      }
                    />
                  </td>

                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Количество"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "Количество", e.target.value)
                      }
                    />
                  </td>
                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Цена"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "Цена", e.target.value)
                      }
                    />
                  </td>
                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Сумма"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "Сумма", e.target.value)
                      }
                    />
                  </td>
                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Дата приобретения"]}
                      onChange={(e) =>
                        handleEdit(
                          rowIndex,
                          "Дата приобретения",
                          e.target.value
                        )
                      }
                    />
                  </td>
                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Поставщик"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "Поставщик", e.target.value)
                      }
                    />
                  </td>

                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["ИНН"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "ИНН", e.target.value)
                      }
                    />
                  </td>

                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Примечание"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "Примечание", e.target.value)
                      }
                    />
                  </td>
                  <td style={cellStyle}>
                    <input
                      style={inputStyle}
                      type="text"
                      value={row["Предложение"]}
                      onChange={(e) =>
                        handleEdit(rowIndex, "Предложение", e.target.value)
                      }
                    />
                  </td>
                </tr>
              </React.Fragment>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

export default ExcelImportExport;
