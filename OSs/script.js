function goHome() {
    window.location.href = './home.html';
}

// Função para buscar e exibir dados da planilha
function fetchOSData() {
    const url = 'https://script.google.com/macros/s/AKfycbwYBlJsP30vtZ7L4FCDwXi6pf5hRAt8CVswdCEL2hl-t7UZYE0Ra18qxTxJwFYgA5QByA/exec';

    fetch(url)
        .then(response => response.json())
        .then(data => {
            const tableBody = document.querySelector('#itemsTable tbody');
            tableBody.innerHTML = ''; // Limpar qualquer dado anterior

            data.forEach(item => {
                const row = document.createElement('tr');
                
                row.innerHTML = `
                    <td>${item.numero_os}</td>
                    <td>${item.data_abertura}</td>
                    <td>${item.data_fechamento}</td>
                    <td>${item.solicitante}</td>
                    <td>${item.departamento}</td>
                    <td>${item.descricao_problema}</td>
                    <td>${item.tecnico}</td>
                    <td>${item.status}</td>
                    <td>${item.observacoes}</td>
                `;
                
                tableBody.appendChild(row);
            });
        })
        .catch(error => console.error('Erro ao buscar dados:', error));
}

// Função para pesquisar itens com base na entrada do usuário
function searchItems() {
    const searchInput = document.getElementById('searchInput').value.toLowerCase();
    const searchOption = document.getElementById('searchOptions').value;
    const rows = document.querySelectorAll('#itemsTable tbody tr');

    rows.forEach(row => {
        const cellValue = row.querySelector(`td:nth-child(${getColumnIndex(searchOption)})`).innerText.toLowerCase();
        if (cellValue.includes(searchInput)) {
            row.style.display = ''; // Mostra a linha
        } else {
            row.style.display = 'none'; // Esconde a linha
        }
    });
}

// Função auxiliar para determinar o índice da coluna com base na opção selecionada
function getColumnIndex(option) {
    switch (option) {
        case 'numero_os': return 1;
        case 'Descricao': return 6;
        case 'Tipo': return 7; // Se houver uma coluna de "Tipo", ajuste conforme necessário
        case 'DataCadastro': return 2;
        case 'DataVencimento': return 3;
        default: return 1;
    }
}

// Função para exportar a tabela para um arquivo Excel
function exportToExcel() {
    const table = document.querySelector('#itemsTable');
    const wb = XLSX.utils.table_to_book(table, { sheet: "Sheet1" });
    XLSX.writeFile(wb, 'Lista_de_OS.xlsx');
    console.log('Exportando para Excel');
}

// Chama a função ao carregar a página para buscar dados da planilha
window.onload = fetchOSData;
