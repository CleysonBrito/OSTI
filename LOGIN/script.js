document.getElementById('loginForm').addEventListener('submit', function(event) {
    event.preventDefault();
    
    const email = document.getElementById('email').value;
    const password = document.getElementById('password').value;
    const errorMessage = document.getElementById('errorMessage');
    
    if ((email === 'cleyson@reciclar.org.br' && password === '@123abc@') ||
        (email === 'danilo.manzoli@reciclar.org.br' && password === 'reciclar1234')) {
        alert('Login bem-sucedido!');
        errorMessage.textContent = '';
        window.location.href = 'cadasto_os.html'; // Redireciona para home.html
    } else {
        errorMessage.textContent = 'E-mail ou senha incorretos.';
    }
});






