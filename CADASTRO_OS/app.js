document.getElementById('dataForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const sku = document.getElementById('sku').value;
    const descricao = document.getElementById('descricao').value;
    const tipo = document.getElementById('tipo').value;
    const unidade = document.getElementById('unidade').value;
    const grupo = document.getElementById('grupo').value;
    const quantidade = document.getElementById('quantidade').value;
    const valor_unitario = document.getElementById('valor_unitario').value;
    const valor_total = document.getElementById('valor_total').value;
    const fornecedor = document.getElementById('fornecedor').value;
    const data_cadastro = document.getElementById('data_cadastro').value;
    const data_vencimento = document.getElementById('data_vencimento').value;

    // Salvar os dados no Excel
    salvarNoExcel({
        sku,
        descricao,
        tipo,
        unidade,
        grupo,
        quantidade,
        valor_unitario,
        valor_total,
        fornecedor,
        data_cadastro,
        data_vencimento
    });

    alert('Produto salvo com sucesso!');
    document.getElementById('dataForm').reset();
});

function calcularValorTotal() {
    const quantidade = document.getElementById('quantidade').value;
    const valor_unitario = document.getElementById('valor_unitario').value;
    const valor_total = quantidade * valor_unitario;
    document.getElementById('valor_total').value = valor_total.toFixed(2);
}

function limparFormulario() {
    document.getElementById('dataForm').reset();
}

function irParaHome() {
    window.location.href = "./home.html"; // Substitua pelo caminho correto para a página inicial
}

function salvarNoExcel(dados) {
    const filePath = 'D:/Sistema Cozinha/HTML JAVA E CSS/SERVIDOR/Banco.xlsx';
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet([dados]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Cadastro');
    XLSX.writeFile(workbook, filePath);
}
