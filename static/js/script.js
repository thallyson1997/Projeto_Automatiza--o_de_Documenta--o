// Aguarda o carregamento do DOM
document.addEventListener('DOMContentLoaded', function() {
    const floatingBtn = document.getElementById('floatingBtn');
    const formsContainer = document.getElementById('formsContainer');
    const submitBtn = document.getElementById('submitBtn');
    const previewBtn = document.getElementById('previewBtn');
    const submitUrl = document.body.dataset.submitUrl || '/gerar-documento';
    const previewArea = document.getElementById('previewArea');
    const previewUrl = '/preview-documento';
    const previewPdfUrl = '/preview-documento-pdf';
    
    // Monta FormData com todos os formularios, incluindo imagens.
    function buildFormsFormData() {
        const forms = document.querySelectorAll('.document-form');
        const formDataGeral = new FormData();

        forms.forEach((form, index) => {
            const unidade = form.querySelector('textarea[name="unidade"]').value;
            const data = form.querySelector('input[name="data"]').value;
            const legenda = form.querySelector('textarea[name="legenda"]').value;
            const imagensInput = form.querySelector('input[name="imagens"]');

            formDataGeral.append(`unidade-${index}`, unidade);
            formDataGeral.append(`data-${index}`, data);
            formDataGeral.append(`legenda-${index}`, legenda);

            if (imagensInput && imagensInput.files) {
                for (let i = 0; i < imagensInput.files.length; i++) {
                    formDataGeral.append(`imagens-${index}`, imagensInput.files[i]);
                }
            }
        });

        return formDataGeral;
    }

        // Pré-visualização PDF
        if (previewBtn && previewArea) {
            previewBtn.addEventListener('click', function(e) {
                e.preventDefault();
                const formDataPreview = buildFormsFormData();
                previewBtn.disabled = true;
                previewBtn.textContent = 'Pré-visualizando...';
                fetch(previewPdfUrl, {
                    method: 'POST',
                    body: formDataPreview
                })
                .then(response => response.json())
                .then(data => {
                    if (data.pdf_url) {
                        previewArea.innerHTML = `<iframe src="${data.pdf_url}#toolbar=0" style="width:100%;height:500px;border:none;"></iframe>`;
                    } else {
                        previewArea.innerHTML = `<div class="preview-placeholder">${data.erro || 'Erro ao gerar PDF.'}</div>`;
                    }
                    previewBtn.disabled = false;
                    previewBtn.textContent = 'Pré-visualizar';
                })
                .catch(error => {
                    previewArea.innerHTML = '<div class="preview-placeholder">Erro ao gerar pré-visualização PDF.</div>';
                    previewBtn.disabled = false;
                    previewBtn.textContent = 'Pré-visualizar';
                });
            });
        }
    
    let formCount = 1;
    let addingForm = false;
    
    // Função para mostrar notificação toast
    function showToast(message, type = 'success', duration = 4000) {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        
        const icons = {
            success: '✓',
            error: '✕',
            info: 'ℹ'
        };
        
        toast.innerHTML = `
            <span class="toast-icon">${icons[type]}</span>
            <span>${message}</span>
        `;
        
        document.body.appendChild(toast);
        
        setTimeout(() => {
            toast.classList.add('hide');
            setTimeout(() => {
                toast.remove();
            }, 400);
        }, duration);
    }
    
    // Função para criar novo formulário
    function addNewForm() {
        if (addingForm) return; // Evita duplicação
        addingForm = true;
        
        formCount++;
        const formIndex = formCount - 1;
        
        const newFormWrapper = document.createElement('div');
        newFormWrapper.className = 'form-wrapper';
        newFormWrapper.setAttribute('data-form-number', formCount);
        
        newFormWrapper.innerHTML = `
            <div class="form-header">
                <h3>Formulário ${formCount}</h3>
                <button type="button" class="btn-remove-form">×</button>
            </div>
            <form class="document-form" data-form-index="${formIndex}">
                <div class="form-row">
                    <div class="form-group form-group-small">
                        <label for="unidade-${formIndex}">Unidade:</label>
                        <textarea id="unidade-${formIndex}" name="unidade" placeholder="Digite a unidade..." rows="1" required></textarea>
                    </div>

                    <div class="form-group">
                        <label for="data-${formIndex}">Data:</label>
                        <input type="date" id="data-${formIndex}" name="data" class="input-date" required>
                    </div>
                </div>

                <div class="form-group">
                    <label for="legenda-${formIndex}">Legenda:</label>
                    <textarea id="legenda-${formIndex}" name="legenda" placeholder="Digite aqui a legenda da documentação..." rows="3" required></textarea>
                </div>

                <div class="form-group">
                    <label for="imagens-${formIndex}">Imagens (até 4):</label>
                    <input type="file" id="imagens-${formIndex}" name="imagens" accept="image/*" multiple required>
                    <p class="form-hint">Selecione até 4 imagens. Formatos aceitos: JPG, PNG, GIF, WebP</p>
                </div>

                <div class="image-preview-container" style="display: none;">
                    <label>Imagens Carregadas:</label>
                    <div class="image-grid"></div>
                </div>
            </form>
        `;
        
        formsContainer.appendChild(newFormWrapper);
        
        // Adicionar listeners
        setupFormListeners(newFormWrapper);
        setupRemoveButton(newFormWrapper);
        
        // Sincroniza se checkboxes estão marcadas
        syncRepeatFieldsIfNeeded();
        
        addingForm = false;
    }
    
    // Setup do botão remover
    function setupRemoveButton(formWrapper) {
        const removeBtn = formWrapper.querySelector('.btn-remove-form');
        removeBtn.addEventListener('click', function(e) {
            e.preventDefault();
            formWrapper.style.animation = 'slideOut 0.3s ease';
            setTimeout(() => {
                formWrapper.remove();
                formCount--;
                
                // Renumerar
                document.querySelectorAll('.form-wrapper').forEach((wrapper, index) => {
                    wrapper.querySelector('.form-header h3').textContent = `Formulário ${index + 1}`;
                });
            }, 300);
        });
    }
    
    // Botão flutuante
    if (floatingBtn) {
        floatingBtn.addEventListener('click', function(e) {
            e.preventDefault();
            addNewForm();
        });
    }
    
    // Setup de listeners de imagens
    function setupFormListeners(formWrapper) {
        const imagensInput = formWrapper.querySelector('input[name="imagens"]');
        const previewContainer = formWrapper.querySelector('.image-preview-container');
        const imageGrid = previewContainer.querySelector('.image-grid');
        let currentFiles = [];
        let draggedIndex = null;
        
        if (!imagensInput) return;

        function syncInputFiles() {
            const dt = new DataTransfer();
            currentFiles.forEach(file => dt.items.add(file));
            imagensInput.files = dt.files;
        }

        function renderImageGrid() {
            imageGrid.innerHTML = '';

            if (currentFiles.length === 0) {
                previewContainer.style.display = 'none';
                return;
            }

            currentFiles.forEach((file, index) => {
                const reader = new FileReader();

                reader.onload = function(event) {
                    const preview = document.createElement('div');
                    preview.className = 'image-preview';
                    preview.setAttribute('draggable', 'true');
                    preview.dataset.index = index;
                    preview.innerHTML = `
                        <img src="${event.target.result}" alt="Prévia ${index + 1}">
                        <button type="button" class="image-preview-remove" data-index="${index}">×</button>
                    `;

                    preview.addEventListener('dragstart', function() {
                        draggedIndex = parseInt(this.dataset.index, 10);
                        this.classList.add('dragging');
                    });

                    preview.addEventListener('dragend', function() {
                        this.classList.remove('dragging');
                        draggedIndex = null;
                    });

                    preview.addEventListener('dragover', function(e) {
                        e.preventDefault();
                    });

                    preview.addEventListener('drop', function(e) {
                        e.preventDefault();

                        const targetIndex = parseInt(this.dataset.index, 10);
                        if (draggedIndex === null || draggedIndex === targetIndex) {
                            return;
                        }

                        const movedFile = currentFiles[draggedIndex];
                        currentFiles.splice(draggedIndex, 1);
                        currentFiles.splice(targetIndex, 0, movedFile);

                        syncInputFiles();
                        renderImageGrid();
                    });

                    imageGrid.appendChild(preview);

                    preview.querySelector('.image-preview-remove').addEventListener('click', function(e) {
                        e.preventDefault();
                        const removeIndex = parseInt(this.dataset.index, 10);
                        currentFiles.splice(removeIndex, 1);
                        syncInputFiles();
                        renderImageGrid();
                    });
                };

                reader.readAsDataURL(file);
            });

            previewContainer.style.display = 'block';
        }
        
        imagensInput.addEventListener('change', function(e) {
            currentFiles = Array.from(this.files).slice(0, 4);

            // Mantem o input sincronizado com o limite de 4 imagens.
            syncInputFiles();
            renderImageGrid();

            if (this.files.length > 4) {
                showToast('Foram consideradas apenas as 4 primeiras imagens.', 'info');
            }
        });
    }
    
    // Setup de checkboxes de repetição (apenas no primeiro formulário)
    function setupRepeatCheckboxes() {
        const repeatUnidadeCheckbox = document.getElementById('repeat-unidade-0');
        const repeatDataCheckbox = document.getElementById('repeat-data-0');
        const unidadeInput = document.getElementById('unidade-0');
        const dataInput = document.getElementById('data-0');
        
        if (!repeatUnidadeCheckbox || !unidadeInput) return;
        
        // Sincroniza Unidade
        function syncUnidade() {
            const unidadeValue = unidadeInput.value;
            const otherForms = document.querySelectorAll('.document-form:not([data-form-index="0"])');
            
            otherForms.forEach(form => {
                const formIndex = form.getAttribute('data-form-index');
                const otherUnidadeInput = document.getElementById(`unidade-${formIndex}`);
                
                if (otherUnidadeInput) {
                    if (repeatUnidadeCheckbox.checked) {
                        otherUnidadeInput.value = unidadeValue;
                    } else {
                        otherUnidadeInput.value = '';
                    }
                }
            });
        }
        
        // Sincroniza Data
        function syncData() {
            const dataValue = dataInput.value;
            const otherForms = document.querySelectorAll('.document-form:not([data-form-index="0"])');
            
            otherForms.forEach(form => {
                const formIndex = form.getAttribute('data-form-index');
                const otherDataInput = document.getElementById(`data-${formIndex}`);
                
                if (otherDataInput) {
                    if (repeatDataCheckbox.checked) {
                        otherDataInput.value = dataValue;
                    } else {
                        otherDataInput.value = '';
                    }
                }
            });
        }
        
        // Listeners para checkboxes
        repeatUnidadeCheckbox.addEventListener('change', syncUnidade);
        repeatDataCheckbox.addEventListener('change', syncData);
        
        // Listeners para mudanças nos inputs do primeiro formulário
        unidadeInput.addEventListener('input', () => {
            if (repeatUnidadeCheckbox.checked) {
                syncUnidade();
            }
        });
        
        dataInput.addEventListener('change', () => {
            if (repeatDataCheckbox.checked) {
                syncData();
            }
        });
    }
    
    // Sincroniza campos se checkboxes estão marcadas (usado ao criar novo formulário)
    function syncRepeatFieldsIfNeeded() {
        const repeatUnidadeCheckbox = document.getElementById('repeat-unidade-0');
        const repeatDataCheckbox = document.getElementById('repeat-data-0');
        
        if (!repeatUnidadeCheckbox) return;
        
        if (repeatUnidadeCheckbox.checked) {
            const unidadeInput = document.getElementById('unidade-0');
            const unidadeValue = unidadeInput.value;
            const otherForms = document.querySelectorAll('.document-form:not([data-form-index="0"])');
            
            otherForms.forEach(form => {
                const formIndex = form.getAttribute('data-form-index');
                const otherUnidadeInput = document.getElementById(`unidade-${formIndex}`);
                if (otherUnidadeInput) {
                    otherUnidadeInput.value = unidadeValue;
                }
            });
        }
        
        if (repeatDataCheckbox && repeatDataCheckbox.checked) {
            const dataInput = document.getElementById('data-0');
            const dataValue = dataInput.value;
            const otherForms = document.querySelectorAll('.document-form:not([data-form-index="0"])');
            
            otherForms.forEach(form => {
                const formIndex = form.getAttribute('data-form-index');
                const otherDataInput = document.getElementById(`data-${formIndex}`);
                if (otherDataInput) {
                    otherDataInput.value = dataValue;
                }
            });
        }
    }
    
    // Setup primeiro formulário
    const firstFormWrapper = document.querySelector('.form-wrapper');
    if (firstFormWrapper) {
        setupFormListeners(firstFormWrapper);
        setupRemoveButton(firstFormWrapper);
        setupRepeatCheckboxes();
    }
    
    // Submit
    if (submitBtn) {
        submitBtn.addEventListener('click', function(e) {
            e.preventDefault();
            
            const forms = document.querySelectorAll('.document-form');
            
            if (forms.length === 0) {
                showToast('Adicione pelo menos um formulário', 'error');
                return;
            }
            
            // Cria FormData com todos os formulários
            const formDataGeral = buildFormsFormData();
            
            submitBtn.disabled = true;
            submitBtn.textContent = 'Gerando...';
            
            // Envia um único request com todos os formulários
            fetch(submitUrl, {
                method: 'POST',
                body: formDataGeral
            })
            .then(response => {
                if (response.ok) {
                    return response.blob().then(blob => {
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = 'documentacao.docx';
                        document.body.appendChild(a);
                        a.click();
                        window.URL.revokeObjectURL(url);
                        document.body.removeChild(a);
                        
                        showToast(`Documento gerado com sucesso! ${forms.length} página(s) criada(s) 🎉`, 'success', 5000);
                        
                        // Limpa todos os formulários
                        document.querySelectorAll('.form-wrapper').forEach((wrapper, idx) => {
                            if (idx > 0) wrapper.remove();
                            else {
                                wrapper.querySelectorAll('input, textarea').forEach(input => {
                                    input.value = '';
                                });
                                wrapper.querySelector('.image-preview-container').style.display = 'none';
                                wrapper.querySelector('.image-grid').innerHTML = '';
                            }
                        });
                        formCount = 1;
                        submitBtn.disabled = false;
                        submitBtn.textContent = 'Gerar Documento';
                    });
                } else {
                    return response.json().then(data => {
                        throw new Error(data.erro || 'Erro ao gerar documento');
                    });
                }
            })
            .catch(error => {
                console.error('Erro:', error);
                showToast(`Erro ao gerar documento: ${error.message}`, 'error', 5000);
                submitBtn.disabled = false;
                submitBtn.textContent = 'Gerar Documento';
            });
        });
    }
});
