// Aguarda o carregamento do DOM
document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('documentForm');
    
    if (form) {
        form.addEventListener('submit', function(e) {
            e.preventDefault();
            
            // Obtém os dados do formulário
            const formData = new FormData(form);
            
            // Desabilita o botão de submit
            const submitBtn = form.querySelector('button[type="submit"]');
            submitBtn.disabled = true;
            submitBtn.textContent = 'Gerando...';
            
            // Faz a requisição
            fetch('/gerar-documento', {
                method: 'POST',
                body: formData
            })
            .then(response => {
                if (response.ok) {
                    // Se sucesso, faz o download do arquivo
                    return response.blob().then(blob => {
                        // Cria um link temporário para download
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = 'documento.docx';
                        document.body.appendChild(a);
                        a.click();
                        window.URL.revokeObjectURL(url);
                        document.body.removeChild(a);
                        
                        // Limpa o formulário
                        form.reset();
                        alert('Documento gerado e baixado com sucesso!');
                    });
                } else {
                    // Se erro, tenta ler a resposta de erro
                    return response.json().then(data => {
                        throw new Error(data.erro || 'Erro ao gerar documento');
                    });
                }
            })
            .catch(error => {
                console.error('Erro:', error);
                alert('Erro ao gerar documento: ' + error.message);
            })
            .finally(() => {
                // Reabilita o botão
                submitBtn.disabled = false;
                submitBtn.textContent = 'Gerar Documento';
            });
        });
    }
});
