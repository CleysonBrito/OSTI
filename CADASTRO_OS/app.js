document.getElementById('dataForm').addEventListener('submit', (e) => {
    e.preventDefault();
    
    const data = {
        numero_os: document.getElementById('numero_os').value,
        data_abertura: document.getElementById('data_abertura').value,
        data_vencimento: document.getElementById('data_vencimento').value,
        solicitante: document.getElementById('solicitante').value,
        departamento: document.getElementById('departamento').value,
        descricao_problema: document.getElementById('descricao_problema').value, // Removido acento
        tecnico: document.getElementById('tecnico').value,
        status: document.getElementById('status').value,
        observacoes: document.getElementById('observacoes').value // Removido acento
    };

    fetch('https://script.google.com/macros/s/AKfycbwYBlJsP30vtZ7L4FCDwXi6pf5hRAt8CVswdCEL2hl-t7UZYE0Ra18qxTxJwFYgA5QByA/exec', {
        method: 'POST',
        mode: 'no-cors',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
    })
    .then(() => {
        alert('OS salva com sucesso!');
        document.getElementById('dataForm').reset();
    })
    .catch(error => console.error('Erro:', error));
});

// Removido o código duplicado relacionado à salvar no Excel

function limparFormulario() {
    document.getElementById('dataForm').reset();
}

function irParaHome() {
    window.location.href = "./home.html"; // Substitua pelo caminho correto para a página inicial
}
