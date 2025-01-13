// Variáveis globais
let dataList = [];

// Função para carregar o arquivo Excel
document.getElementById('excelFile').addEventListener('change', (event) => {
  const file = event.target.files[0];
  if (file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      dataList = XLSX.utils.sheet_to_json(sheet, { header: 1 }).flat();
      
      // Habilitar a barra de busca
      document.getElementById('searchBar').disabled = false;
      alert('Arquivo Excel carregado com sucesso!');
    };
    reader.readAsArrayBuffer(file);
  }
});

// Função de busca
document.getElementById('searchBar').addEventListener('input', (event) => {
  const searchQuery = event.target.value.toLowerCase();
  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = '';

  if (searchQuery) {
    const filteredResults = dataList.filter((item) =>
      item.toString().toLowerCase().includes(searchQuery)
    );

    if (filteredResults.length > 0) {
      filteredResults.forEach((result) => {
        const p = document.createElement('p');
        p.textContent = result;
        resultsDiv.appendChild(p);
      });
    } else {
      resultsDiv.innerHTML = '<p>Nenhum resultado encontrado.</p>';
    }
  }
});
