document.addEventListener("DOMContentLoaded", function () {
  const fileInput = document.getElementById("file-input");
  const processButton = document.getElementById("process-button");
  processButton.addEventListener("click", function () {
    const file = fileInput.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Realiza o processamento e formatação dos dados
        jsonData.forEach(function (row) {
          // Formata todas as células da terceira coluna com duas casas decimais separadas por vírgula
          if (typeof row[2] === "number") {
            row[2] = row[2].toFixed(2).replace(".", ",");
          } else if (typeof row[2] === "string") {
            row[2] = row[2].replace(".", ",");
          }
          // Garante que os dados da primeira coluna contenham 000000 dígitos
          row[0] = ("000000" + row[0]).slice(-6);
        });

        // Cria um novo workbook com os dados processados
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.aoa_to_sheet(jsonData);

        // Centraliza os dados em todas as células do novo worksheet
        for (const key in newWorksheet) {
          if (key !== "!ref" && newWorksheet.hasOwnProperty(key)) {
            newWorksheet[key].s = { alignment: { horizontal: "center" } };
          }
        }

        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");

        // Converte o novo workbook em um arquivo Excel e inicia o download
        const newFileData = XLSX.write(newWorkbook, {
          bookType: "xlsx",
          type: "array",
        });
        const blob = new Blob([new Uint8Array(newFileData)], {
          type: "application/octet-stream",
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "Formatado_Corretamente.xlsx";
        a.click();

        console.log("Arquivo formatado e salvo com sucesso!");
      };
      reader.readAsArrayBuffer(file);
    } else {
      console.log("Nenhum arquivo selecionado.");
    }
  });
});
