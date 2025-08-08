// COLOQUE A URL DO SEU NOVO WEB APP DO APPS SCRIPT AQUI!
// Será algo como: https://script.google.com/macros/s/SEU_ID_DE_DEPLOYMENT/exec
const APPS_SCRIPT_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbz0mrV_qsAbgj6kOb8L23-rcgDQzwOoxLfKc18wEyPWZDhZkk3NHruZv98_PwaGjQSQjA/exec'; // <--- ATUALIZE ESTA URL SE NECESSÁRIO!

let currentProfile = 'separacao'; // Perfil padrão ao carregar
let data = []; // Armazena todos os dados do Google Sheet
let filteredData = []; // Armazena os dados após a aplicação dos filtros
let currentEditItem = null; // Item sendo editado no modal
let refreshIntervalId; // ID para controlar o auto-refresh

// --- Funções de UI e Feedback ---

/**
 * Exibe o overlay de carregamento (spinner).
 */
function showLoading() {
    document.getElementById('loading-overlay').classList.add('visible');
}

/**
 * Oculta o overlay de carregamento (spinner).
 */
function hideLoading() {
    document.getElementById('loading-overlay').classList.remove('visible');
}

let currentAlertTimeout;
/**
 * Exibe um alerta temporário na interface do usuário.
 * @param {string} message - A mensagem a ser exibida.
 * @param {'success' | 'danger' | 'info'} type - O tipo de alerta (para estilização).
 */
function showAlert(message, type) {
    const alertContainer = document.getElementById('alert-container');
    alertContainer.innerHTML = `<div class="alert alert-${type}">${message}</div>`;
    alertContainer.style.display = 'block';

    if (currentAlertTimeout) {
        clearTimeout(currentAlertTimeout);
    }
    currentAlertTimeout = setTimeout(() => {
        alertContainer.style.display = 'none';
        alertContainer.innerHTML = '';
    }, 5000); // Alerta desaparece após 5 segundos
    console.log(`Alerta (${type}): ${message}`);
}

// --- Funções de Inicialização e Controle de Perfil ---

document.addEventListener('DOMContentLoaded', function() {
    // Esconder todos os modais no carregamento inicial para evitar que apareçam
    document.getElementById('edit-modal').style.display = 'none';
    document.getElementById('modal-to-pca').style.display = 'none';
    document.getElementById('modal-from-pca').style.display = 'none';

    setupFilters();
    loadDataFromGoogleSheet(); // Inicia o carregamento de dados
    document.getElementById('btn-separacao').classList.add('active'); // Define Separação como perfil inicial
    
    // ATENÇÃO: Chamar togglePcaColumnsVisibility aqui para definir a visibilidade inicial
    togglePcaColumnsVisibility(currentProfile); 

    // Inicia a atualização automática a cada 30 segundos
    startAutoRefresh(30000); // 30000 milissegundos = 30 segundos

    // Listener para o card de Clientes Únicos (para o tooltip)
    const clientesUniqueCard = document.getElementById('clientes-unique-card');
    const clientesTooltip = document.getElementById('clientes-tooltip');

    if (!clientesUniqueCard || !clientesTooltip) {
        console.error('Tooltip: Elementos clientesUniqueCard ou clientesTooltip não encontrados no DOM. Verifique seu HTML.');
    }

    // Event listeners para os modais de ação
    document.getElementById('form-to-pca').addEventListener('submit', handleSendToPca);
    document.getElementById('form-from-pca').addEventListener('submit', handleReceiveFromPca);
    
    addFilterListeners(); // Adiciona os listeners para os filtros
});

/**
 * Inicia a atualização automática dos dados em um intervalo definido.
 * @param {number} intervalTime - Tempo em milissegundos entre as atualizações.
 */
function startAutoRefresh(intervalTime) {
    if (refreshIntervalId) {
        clearInterval(refreshIntervalId); // Limpa qualquer intervalo existente
    }
    refreshIntervalId = setInterval(() => {
        console.log('Auto-refresh: Puxando dados do Google Sheet...');
        loadDataFromGoogleSheet();
    }, intervalTime);
    console.log(`Auto-refresh iniciado a cada ${intervalTime / 1000} segundos.`);
}

/**
 * Altera o perfil de usuário e atualiza a interface.
 * @param {string} profile - O perfil a ser ativado ('separacao', 'embalagem', 'compras', 'pca', 'admin').
 */
function changeProfile(profile) {
    console.log('changeProfile: Alterando para o perfil:', profile);
    currentProfile = profile;

    // Remove 'active' de todos os botões e adiciona ao botão clicado
    document.querySelectorAll('.profile-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    document.getElementById(`btn-${profile}`).classList.add('active');

    // Resetar filtros específicos que podem não se aplicar a todos os perfis
    document.getElementById('filter-status').value = '';
    document.getElementById('filter-responsavel-pca').value = '';
    document.getElementById('filter-tipo-servico-pca').value = '';

    togglePcaColumnsVisibility(profile); // Atualiza a visibilidade das colunas PCA
    updateDisplay(); // Atualiza a tabela e estatísticas
}

/**
 * Alterna a visibilidade das colunas e filtros relacionados a PCA.
 * @param {string} profile - O perfil atual.
 */
function togglePcaColumnsVisibility(profile) {
    const pcaColumns = document.querySelectorAll('.pca-column');
    const displayPcaInfo = (profile === 'pca' || profile === 'admin'); // Define se as colunas PCA devem ser visíveis

    pcaColumns.forEach(col => {
        // Para os THs e TDs da tabela, use 'table-cell'
        if (col.tagName === 'TH' || col.tagName === 'TD') {
            col.style.display = displayPcaInfo ? 'table-cell' : 'none';
        } 
        // Para os inputs de filtro, use 'block' ou 'flex' conforme seu layout de filtros
        else if (col.classList.contains('filter-input')) {
            col.style.display = displayPcaInfo ? 'block' : 'none';
        }
    });
    console.log(`togglePcaColumnsVisibility: Colunas PCA ${displayPcaInfo ? 'visíveis' : 'ocultas'} para o perfil: ${profile}`);
}


// --- Funções de Comunicação com Google Apps Script ---

/**
 * Carrega os dados da planilha do Google Sheet via Apps Script.
 */
async function loadDataFromGoogleSheet() {
    console.log('loadDataFromGoogleSheet() chamada. Iniciando requisição GET...');
    showLoading();
    try {
        const response = await fetch(APPS_SCRIPT_WEB_APP_URL, {
            method: 'GET',
            headers: { 'Content-Type': 'application/json' }
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Erro na rede ou no servidor Apps Script: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const result = await response.json();

        if (result.error) {
            throw new Error(`Erro do Apps Script: ${result.error} - ${result.details || ''}`);
        }

        // Mapeia o resultado para adicionar a propriedade 'selected' e garantir o formato
        data = result.map(item => ({
            ...item,
            selected: item.selected || false, // Para seleção em massa
            dataPedido: item.dataPedido || '', 
            dataEnvio: item.dataEnvio || '',
            previsaoRetorno: item.previsaoRetorno || '',
            dataRetorno: item.dataRetorno || ''
        }));

        console.log('loadDataFromGoogleSheet: Dados recarregados. Total de itens:', data.length);
        updateDisplay(); // Atualiza a interface com os novos dados

    } catch (error) {
        showAlert(`Erro ao carregar dados: ${error.message}`, 'danger');
        console.error('Erro ao carregar dados do Google Sheet:', error);
    } finally {
        hideLoading();
    }
}

/**
 * Envia dados para o Google Sheet (atualização).
 * @param {object|Array<object>} itemOrItemsToUpdate - Um único item ou um array de itens para atualizar.
 * Cada item deve conter _rowIndex para atualização.
 */
async function updateItemsInGoogleSheet(itemOrItemsToUpdate) {
    console.log(`updateItemsInGoogleSheet() chamada. Dados a enviar:`, itemOrItemsToUpdate);
    showLoading();
    try {
        const payload = {
            action: 'UPDATE', // Sempre 'UPDATE' agora
            data: itemOrItemsToUpdate
        };

        const response = await fetch(APPS_SCRIPT_WEB_APP_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Erro na rede ou no servidor Apps Script ao atualizar: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const result = await response.json();

        if (result.error) {
            throw new Error(`Erro do Apps Script ao atualizar: ${result.error} - ${result.details || ''}`);
        }

        // Não mostra showAlert aqui, pois as funções de modal/ação farão isso.
        console.log(`Operação de atualização bem-sucedida. Recarregando dados...`);
        await loadDataFromGoogleSheet(); // Recarrega os dados para refletir as mudanças

    } catch (error) {
        // Lança o erro para que a função chamadora (handleSendToPca, etc.) possa capturá-lo
        throw error; 
    } finally {
        hideLoading();
    }
}


// --- Lógica de Dados e Status ---

/**
 * Mapeia o status interno para um rótulo amigável.
 * @param {string} status - O status interno (camelCase).
 * @returns {string} O rótulo formatado.
 */
function getStatusLabel(status) {
    const labels = {
        'emSeparacao': 'Separação',
        'separado': 'Separado',
        'embalagem': 'Embalagem',
        'emCompras': 'Compras',
        'emPca': 'PCA',
        'retornoPca': 'Retorno PCA',
        'concluido': 'Concluído'
    };
    return labels[status] || status; // Retorna o status original se não estiver mapeado (fallback)
}

// --- Atualização da Interface ---

/**
 * Atualiza todos os elementos da interface (filtros, tabela, estatísticas).
 */
function updateDisplay() {
    console.log('updateDisplay() chamada. Aplicando filtros e renderizando...');
    setupFilters(); // Garante que os dropdowns de filtro estejam atualizados com base nos dados
    filterData(); // Primeiro filtra
    renderTable(); // Depois renderiza a tabela
    updateStats(); // Atualiza os números de estatísticas
    updateNotificationBadges(); // Atualiza os contadores nos botões de perfil
    console.log('Display atualizado. Itens filtrados para exibição:', filteredData.length);
}

/**
 * Preenche os dropdowns de filtro com valores únicos dos dados.
 */
function setupFilters() {
    const uniqueResponsaveis = new Set();
    const uniqueTiposServico = new Set();

    data.forEach(item => {
        if (item.responsavelPca) uniqueResponsaveis.add(item.responsavelPca);
        if (item.tipoServicoPca) uniqueTiposServico.add(item.tipoServicoPca);
    });

    populateFilterDropdown('filter-responsavel-pca', Array.from(uniqueResponsaveis).sort());
    populateFilterDropdown('filter-tipo-servico-pca', Array.from(uniqueTiposServico).sort());
}

/**
 * Preenche um elemento select com opções.
 * @param {string} elementId - O ID do elemento select.
 * @param {Array<string>} options - Um array de strings para as opções.
 */
function populateFilterDropdown(elementId, options) {
    const select = document.getElementById(elementId);
    if (!select) return;

    // Salva o valor selecionado antes de limpar
    const currentValue = select.value;

    select.innerHTML = '<option value="">Todos</option>'; // Opção padrão
    options.forEach(option => {
        const opt = document.createElement('option');
        opt.value = option;
        opt.textContent = option;
        select.appendChild(opt);
    });

    // Restaura o valor selecionado, se ainda existir nas novas opções
    if ([...options, ''].includes(currentValue)) {
        select.value = currentValue;
    } else {
        select.value = ''; // Reseta se a opção anterior não existe mais
    }
}

/**
 * Adiciona listeners para os inputs de filtro para acionar a filtragem.
 */
function addFilterListeners() {
    document.getElementById('filter-oe').addEventListener('input', updateDisplay);
    document.getElementById('filter-pedido-cliente').addEventListener('input', updateDisplay); // NOVO LISTENER
    document.getElementById('filter-cliente').addEventListener('input', updateDisplay);
    document.getElementById('filter-produto').addEventListener('input', updateDisplay);
    document.getElementById('filter-status').addEventListener('change', updateDisplay);
    document.getElementById('filter-responsavel-pca').addEventListener('change', updateDisplay);
    document.getElementById('filter-tipo-servico-pca').addEventListener('change', updateDisplay);
    console.log('Filter listeners adicionados.');
}

/**
 * Aplica os filtros atuais nos dados brutos e preenche `filteredData`.
 */
function filterData() {
    console.log('filterData() - Perfil atual:', currentProfile);
    let tempFiltered = [...data];

    // Lógica de filtragem baseada no perfil
    switch (currentProfile) {
        case 'separacao':
            tempFiltered = tempFiltered.filter(item => item.statusGlobal === 'emSeparacao');
            break;
        case 'embalagem':
            tempFiltered = tempFiltered.filter(item =>
                item.statusGlobal === 'separado' || item.statusGlobal === 'embalagem' ||
                (item.statusGlobal === 'retornoPca' && item.destinoAposPca === 'embalagem')
            );
            break;
        case 'compras':
            tempFiltered = tempFiltered.filter(item => item.statusGlobal === 'emCompras');
            break;
        case 'pca':
            tempFiltered = tempFiltered.filter(item => item.statusGlobal === 'emPca' || item.statusGlobal === 'retornoPca');
            break;
        case 'admin':
            // Admin vê todos os itens, sem filtragem inicial por status de perfil
            break;
        default:
            tempFiltered = []; // Se o perfil for desconhecido
            break;
    }

    // Filtros de texto e seleção (aplicados sobre os dados já filtrados por perfil)
    const oeFilter = document.getElementById('filter-oe').value.toLowerCase();
    const pedidoClienteFilter = document.getElementById('filter-pedido-cliente').value.toLowerCase(); // NOVO FILTRO
    const clienteFilter = document.getElementById('filter-cliente').value.toLowerCase();
    const produtoFilter = document.getElementById('filter-produto').value.toLowerCase();
    const responsavelPcaFilter = document.getElementById('filter-responsavel-pca').value.toLowerCase();
    const tipoServicoPcaFilter = document.getElementById('filter-tipo-servico-pca').value.toLowerCase();
    const statusFilter = document.getElementById('filter-status').value;

    if (oeFilter) {
        tempFiltered = tempFiltered.filter(item =>
            String(item.oe || '').toLowerCase().includes(oeFilter)
        );
    }
    // LÓGICA DO NOVO FILTRO AQUI
    if (pedidoClienteFilter) {
        tempFiltered = tempFiltered.filter(item =>
            String(item.pedidoCliente || '').toLowerCase().includes(pedidoClienteFilter)
        );
    }
    // FIM DA LÓGICA DO NOVO FILTRO
    if (clienteFilter) {
        tempFiltered = tempFiltered.filter(item =>
            String(item.fantasia || '').toLowerCase().includes(clienteFilter)
        );
    }
    if (produtoFilter) {
        tempFiltered = tempFiltered.filter(item =>
            String(item.produto || '').toLowerCase().includes(produtoFilter)
        );
    }
    if (responsavelPcaFilter) {
        tempFiltered = tempFiltered.filter(item =>
            String(item.responsavelPca || '').toLowerCase().includes(responsavelPcaFilter)
        );
    }
    if (tipoServicoPcaFilter) {
        tempFiltered = tempFiltered.filter(item =>
            String(item.tipoServicoPca || '').toLowerCase().includes(tipoServicoPcaFilter)
        );
    }
    if (statusFilter) {
        tempFiltered = tempFiltered.filter(item => {
            return item.statusGlobal === statusFilter;
        });
    }

    filteredData = tempFiltered;
    console.log('filterData() - Itens após filtragem total:', filteredData.length);
}

/**
 * Atualiza os números exibidos nos cards de estatísticas.
 */
function updateStats() {
    const totalItems = data.length;
    const totalTableItems = filteredData.length;

    const uniqueClientes = new Set();
    const uniquePedidos = new Set();
    data.forEach(item => {
        if (item.fantasia) {
            uniqueClientes.add(item.fantasia);
        }
        if (item.pedidoCliente) {
            uniquePedidos.add(item.pedidoCliente);
        }
    });
    const totalUniqueClientes = uniqueClientes.size;
    const totalUniquePedidos = uniquePedidos.size;

    const emSeparacaoItems = data.filter(item => item.statusGlobal === 'emSeparacao').length;
    const separadoItems = data.filter(item => item.statusGlobal === 'separado').length;
    const embalagemItems = data.filter(item => item.statusGlobal === 'embalagem').length;
    const emComprasItems = data.filter(item => item.statusGlobal === 'emCompras').length;
    const emPcaItems = data.filter(item => item.statusGlobal === 'emPca').length;
    const concluidoItems = data.filter(item => item.statusGlobal === 'concluido').length;

    document.getElementById('total-items').textContent = totalItems;
    document.getElementById('total-table-items').textContent = totalTableItems;
    document.getElementById('total-unique-clientes').textContent = totalUniqueClientes;
    document.getElementById('total-unique-pedidos').textContent = totalUniquePedidos;
    
    document.getElementById('em-separacao-items').textContent = emSeparacaoItems;
    document.getElementById('separado-items').textContent = separadoItems;
    document.getElementById('embalagem-items').textContent = embalagemItems;
    document.getElementById('em-compras-items').textContent = emComprasItems;
    document.getElementById('em-pca-items').textContent = emPcaItems;
    document.getElementById('concluido-items').textContent = concluidoItems;

    // AQUI: Preencher o conteúdo do tooltip de clientes
    const clientesTooltip = document.getElementById('clientes-tooltip');
    if (clientesTooltip) {
        let clientesList = Array.from(uniqueClientes).sort().join('<br>');
        if (clientesList === '') {
            clientesList = 'Nenhum cliente encontrado.';
        }
        clientesTooltip.innerHTML = `<strong>Clientes Únicos:</strong><br>${clientesList}`;
    }

    console.log('updateStats:', { totalItems, totalTableItems, totalUniqueClientes, totalUniquePedidos, emSeparacaoItems, separadoItems, embalagemItems, emComprasItems, emPcaItems, concluidoItems });
}

/**
 * Atualiza os badges de notificação nos botões de perfil.
 */
function updateNotificationBadges() {
    const emSeparacaoCount = data.filter(item => item.statusGlobal === 'emSeparacao').length;
    const separadoCount = data.filter(item => item.statusGlobal === 'separado').length;
    const embalagemCount = data.filter(item => item.statusGlobal === 'embalagem').length; // Contagem para "Embalagem"
    const emComprasCount = data.filter(item => item.statusGlobal === 'emCompras').length;
    const emPcaOrRetornoPcaCount = data.filter(item => item.statusGlobal === 'emPca' || item.statusGlobal === 'retornoPca').length;

    updateBadge(document.getElementById('btn-separacao'), emSeparacaoCount);
    // Para Embalagem, podemos somar 'separado' e 'embalagem' para o badge, se desejado.
    // Ou manter apenas um deles, dependendo do que o usuário de embalagem precisa ver prioritariamente.
    // Por enquanto, vamos somar para mostrar um total de itens que podem ser de interesse da Embalagem.
    updateBadge(document.getElementById('btn-embalagem'), separadoCount + embalagemCount); 
    updateBadge(document.getElementById('btn-compras'), emComprasCount);
    updateBadge(document.getElementById('btn-pca'), emPcaOrRetornoPcaCount);

    console.log('updateNotificationBadges: Em Separação:', emSeparacaoCount, 'Separado:', separadoCount, 'Embalagem:', embalagemCount, 'Em Compras:', emComprasCount, 'Em PCA/Retorno PCA:', emPcaOrRetornoPcaCount);
}

/**
 * Adiciona ou atualiza um badge de notificação em um botão.
 * @param {HTMLElement} buttonElement - O elemento do botão.
 * @param {number} count - O número a ser exibido no badge.
 */
function updateBadge(buttonElement, count) {
    let badge = buttonElement.querySelector('.notification-badge');
    if (count > 0) {
        if (!badge) {
            badge = document.createElement('span');
            badge.className = 'notification-badge';
            buttonElement.appendChild(badge);
        }
        badge.textContent = count;
        badge.classList.remove('hidden'); // Garante que o badge seja visível
    } else {
        if (badge) {
            badge.classList.add('hidden'); // Oculta o badge se o contador for 0
            // badge.remove(); // Ou remove completamente se preferir
        }
    }
}

/**
 * Renderiza a tabela com os dados filtrados.
 */
function renderTable() {
    console.log('renderTable() - Iniciando renderização.');
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    if (filteredData.length === 0) {
        const noDataRow = document.createElement('tr');
        noDataRow.innerHTML = `<td colspan="14" style="text-align: center; padding: 20px;">Nenhum item encontrado para este perfil ou filtros.</td>`;
        tbody.appendChild(noDataRow);
        console.log('renderTable: Nenhum dado filtrado para exibição.');
        return;
    }

    filteredData.forEach((item) => {
        const displayStatusClass = item.statusGlobal;
        const displayStatusLabel = getStatusLabel(item.statusGlobal);

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.oe || ''}</td>
            <td>${item.pedidoCliente || ''}</td> <td>${formatDate(item.dataPedido)}</td>
            <td>${item.fantasia || ''}</td>
            <td>${item.produto || ''}</td>
            <td>${item.modelo || ''}</td>
            <td>${item.qtd || ''}</td>
            <td class="pca-column">${formatDate(item.dataEnvio)}</td>
            <td class="pca-column">${formatDate(item.previsaoRetorno)}</td>
            <td class="pca-column">${formatDate(item.dataRetorno)}</td>
            <td class="pca-column">${item.responsavelPca || ''}</td>
            <td class="pca-column">${item.tipoServicoPca || ''}</td>
            <td><span class="status-badge status-${displayStatusClass}">${displayStatusLabel}</span></td>
            <td>${getActionButtons(item)}</td>
        `;
        tbody.appendChild(row);
    });
    togglePcaColumnsVisibility(currentProfile); 
    console.log('renderTable: Tabela renderizada com', filteredData.length, 'itens.');
}

/**
 * Retorna os botões de ação apropriados para um item com base no perfil e status.
 * @param {object} item - O item de PCA.
 * @returns {string} HTML dos botões de ação.
 */
function getActionButtons(item) {
    let buttons = '';

    if (currentProfile === 'admin') {
        buttons += `<button class="action-btn" onclick="editItem(${item.id})">Editar</button>`;
    } else if (currentProfile === 'separacao') {
        if (item.statusGlobal === 'emSeparacao') {
            buttons += `<button class="action-btn success" onclick="changeStatus(${item.id}, 'separado')">Marcar Separado</button>`;
            buttons += `<button class="action-btn danger" onclick="changeStatus(${item.id}, 'emCompras')">Mover para Compras</button>`;
        }
    } else if (currentProfile === 'embalagem') {
        if (item.statusGlobal === 'separado' || (item.statusGlobal === 'retornoPca' && item.destinoAposPca === 'embalagem')) {
            buttons += `<button class="action-btn" onclick="changeStatus(${item.id}, 'embalagem')">Marcar Embalagem/Máquina</button>`;
            buttons += `<button class="action-btn success" onclick="openToPcaModal(${item.id})">Enviar para PCA</button>`;
            buttons += `<button class="action-btn success" onclick="changeStatus(${item.id}, 'concluido')">Marcar Concluído</button>`;
        } else if (item.statusGlobal === 'embalagem') {
            buttons += `<button class="action-btn success" onclick="changeStatus(${item.id}, 'concluido')">Marcar Concluído</button>`;
            buttons += `<button class="action-btn success" onclick="openToPcaModal(${item.id})">Enviar para PCA</button>`;
        }
    } else if (currentProfile === 'compras') {
        if (item.statusGlobal === 'emCompras') {
            buttons += `<button class="action-btn success" onclick="changeStatus(${item.id}, 'separado')">Item Recebido (Mover para Separado)</button>`;
        }
    } else if (currentProfile === 'pca') {
        if (item.statusGlobal === 'emPca') {
            buttons += `<button class="action-btn success" onclick="openFromPcaModal(${item.id})">Receber do PCA</button>`;
        } else if (item.statusGlobal === 'retornoPca') {
            buttons += `<button class="action-btn success" onclick="openFromPcaModal(${item.id})">Re-apontar Destino</button>`;
        }
    }
    return buttons;
}

/**
 * Formata um valor de data para exibição em pt-BR.
 * @param {*} dateValue - O valor da data (número serial do Excel, string ISO ou objeto Date).
 * @returns {string} Data formatada ou string vazia.
 */
function formatDate(dateValue) {
    if (!dateValue || dateValue === '') return '';

    let date;
    if (typeof dateValue === 'number') {
        date = new Date(Math.round((dateValue - 25569) * 86400 * 1000));
    } else if (typeof dateValue === 'string') {
        // Tenta parsear como ISO ou YYYY-MM-DD
        date = new Date(dateValue);
        // Se a data for inválida após a tentativa direta, tenta adicionar um fuso horário neutro
        if (isNaN(date.getTime())) {
            date = new Date(dateValue + 'T12:00:00'); 
        }
    } else if (dateValue instanceof Date) {
        date = dateValue;
    } else {
        return String(dateValue); 
    }

    if (!isNaN(date.getTime())) {
        return date.toLocaleDateString('pt-BR');
    }
    return String(dateValue);
}

/**
 * Formata um objeto Date para string "YYYY-MM-DD" para campos input type="date".
 * @param {Date|string|number} dateValue - A data a ser formatada.
 * @returns {string} Data formatada como "YYYY-MM-DD" ou string vazia.
 */
function formatToInputDate(dateValue) {
    if (!dateValue || dateValue === '') return '';

    let date;
    if (typeof dateValue === 'number') {
        date = new Date(Math.round((dateValue - 25569) * 86400 * 1000));
    } else if (typeof dateValue === 'string') {
        // Tenta parsear como ISO ou YYYY-MM-DD
        date = new Date(dateValue);
        // Se a data for inválida após a tentativa direta, tenta adicionar um fuso horário neutro
        if (isNaN(date.getTime())) {
            date = new Date(dateValue + 'T12:00:00');
        }
    } else if (dateValue instanceof Date) {
        date = dateValue;
    } else {
        return '';
    }

    if (!isNaN(date.getTime())) {
        const year = date.getFullYear();
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const day = date.getDate().toString().padStart(2, '0');
        return `${year}-${month}-${day}`;
    }
    return '';
}

// --- Lógica de Ações e Modais ---

/**
 * Altera o status global de um item.
 * @param {number} itemId - O ID único do item.
 * @param {string} newStatus - O novo status global (camelCase).
 */
async function changeStatus(itemId, newStatus) {
    const item = data.find(item => item.id === itemId);
    if (!item) {
        showAlert('Erro: Item não encontrado para mudança de status.', 'danger');
        return;
    }

    if (confirm(`Confirmar mudança de status para "${getStatusLabel(newStatus)}" para o item OE: ${item.oe}?`)) {
        const updatedItem = {
            id: item.id,
            _rowIndex: item._rowIndex,
            statusGlobal: newStatus
        };
        // Adicionar lógica de data para o novo status 'separado'
        if (newStatus === 'separado' && !item.dataFimSeparacao) {
            updatedItem.dataFimSeparacao = new Date().toISOString().split('T')[0];
        } else if (newStatus === 'emSeparacao' && !item.dataInicioSeparacao) {
            updatedItem.dataInicioSeparacao = new Date().toISOString().split('T')[0];
        }
        
        try {
            await updateItemsInGoogleSheet(updatedItem);
            showAlert(`Status do item ${item.oe} atualizado para "${getStatusLabel(newStatus)}"!`, 'success');
        } catch (error) {
            showAlert(`Erro ao atualizar status: ${error.message}`, 'danger');
            console.error(`Erro ao atualizar status do item ${item.oe}:`, error);
        }
    }
}

/**
 * Abre o modal de edição para um item específico (apenas para Admin).
 * @param {number} itemId - O ID único do item a ser editado.
 */
function editItem(itemId) {
    const item = data.find(item => item.id === itemId);
    if (!item) {
        console.error('Item não encontrado para edição:', itemId);
        showAlert('Erro: Item não encontrado para edição.', 'danger');
        return;
    }
    currentEditItem = { ...item }; // Cria uma cópia para edição
    console.log('editItem: Abrindo modal para item:', currentEditItem);

    // Preenche os campos do modal
    document.getElementById('edit-oe').value = currentEditItem.oe || '';
    document.getElementById('edit-pedido').value = currentEditItem.pedidoCliente || '';
    document.getElementById('edit-data-pedido').value = formatToInputDate(currentEditItem.dataPedido);
    document.getElementById('edit-fantasia').value = currentEditItem.fantasia || '';
    document.getElementById('edit-produto').value = currentEditItem.produto || '';
    document.getElementById('edit-modelo').value = currentEditItem.modelo || '';
    document.getElementById('edit-qtd').value = currentEditItem.qtd || '';
    
    document.getElementById('edit-data-envio').value = formatToInputDate(currentEditItem.dataEnvio);
    document.getElementById('edit-previsao-retorno').value = formatToInputDate(currentEditItem.previsaoRetorno);
    document.getElementById('edit-data-retorno').value = formatToInputDate(currentEditItem.dataRetorno);
    document.getElementById('edit-responsavel-pca').value = currentEditItem.responsavelPca || '';
    document.getElementById('edit-tipo-servico-pca').value = currentEditItem.tipoServicoPca || '';
    
    // Configura a opção de status global no modal de edição
    const statusSelect = document.getElementById('edit-status-global');
    statusSelect.innerHTML = `
        <option value="emSeparacao">Separação</option>
        <option value="separado">Separado</option> 
        <option value="embalagem">Embalagem</option>
        <option value="emCompras">Compras</option>
        <option value="emPca">PCA</option>
        <option value="retornoPca">Retorno PCA</option>
        <option value="concluido">Concluído</option>
    `;
    statusSelect.value = currentEditItem.statusGlobal; // Define o valor atual

    document.getElementById('edit-modal').style.display = 'flex'; // Usar 'flex' para centralização via CSS
}

/**
 * Lida com o envio do formulário de edição.
 * @param {Event} event - O evento de submit do formulário.
 */
document.getElementById('edit-form').addEventListener('submit', async function(event) {
    event.preventDefault();
    console.log('edit-form: Submit acionado.');

    if (!currentEditItem) {
        showAlert('Erro: Nenhum item selecionado para edição.', 'danger');
        return;
    }

    // Atualiza o objeto currentEditItem com os novos valores do formulário
    currentEditItem.oe = document.getElementById('edit-oe').value.trim();
    currentEditItem.pedidoCliente = document.getElementById('edit-pedido').value.trim();
    currentEditItem.dataPedido = document.getElementById('edit-data-pedido').value; 
    currentEditItem.fantasia = document.getElementById('edit-fantasia').value.trim();
    currentEditItem.produto = document.getElementById('edit-produto').value.trim();
    currentEditItem.modelo = document.getElementById('edit-modelo').value.trim();
    currentEditItem.qtd = parseInt(document.getElementById('edit-qtd').value);
    currentEditItem.dataEnvio = document.getElementById('edit-data-envio').value;
    currentEditItem.previsaoRetorno = document.getElementById('edit-previsao-retorno').value;
    currentEditItem.dataRetorno = document.getElementById('edit-data-retorno').value;
    currentEditItem.responsavelPca = document.getElementById('edit-responsavel-pca').value.trim();
    currentEditItem.tipoServicoPca = document.getElementById('edit-tipo-servico-pca').value.trim();
    currentEditItem.statusGlobal = document.getElementById('edit-status-global').value;
    
    // Validações adicionais
    if (isNaN(currentEditItem.qtd) || currentEditItem.qtd <= 0) {
        showAlert('Quantidade deve ser um número positivo.', 'danger');
        return;
    }
    
    try {
        await updateItemsInGoogleSheet(currentEditItem); // Envia o item completo para atualização
        showAlert('Item atualizado com sucesso!', 'success'); // Alerta de sucesso
        closeModal(); // Fechar o modal após salvar
        console.log('edit-form: Edição enviada para o servidor. Item atualizado:', currentEditItem);
    } catch (error) {
        showAlert(`Erro ao salvar edições: ${error.message}`, 'danger');
        console.error('Erro ao salvar edições:', error);
    }
});

/**
 * Fecha o modal de edição.
 */
function closeModal() {
    document.getElementById('edit-modal').style.display = 'none';
    currentEditItem = null;
    console.log('closeModal: Modal de edição fechado.');
}

/**
 * Abre o modal para enviar item para PCA.
 * @param {number} itemId - ID do item.
 */
function openToPcaModal(itemId) {
    const item = data.find(item => item.id === itemId);
    if (!item) {
        showAlert('Erro: Item não encontrado.', 'danger');
        return;
    }
    document.getElementById('to-pca-item-id').value = item.id;
    document.getElementById('to-pca-oe').textContent = item.oe;
    document.getElementById('to-pca-responsavel').value = item.responsavelPca || '';
    document.getElementById('to-pca-tipo-servico').value = item.tipoServicoPca || '';
    document.getElementById('to-pca-previsao-retorno').value = formatToInputDate(item.previsaoRetorno || new Date()); // Preenche com data atual se vazia
    document.getElementById('modal-to-pca').style.display = 'flex'; // Usar 'flex' para centralização via CSS
}

/**
 * Lida com o envio do formulário de enviar para PCA.
 * @param {Event} event - Evento de submit.
 */
async function handleSendToPca(event) {
    event.preventDefault(); // <-- CRUCIAL: Garante que o formulário não recarregue a página
    console.log("1. handleSendToPca iniciado."); // Log para depuração

    const itemId = parseInt(document.getElementById('to-pca-item-id').value);
    const responsavel = document.getElementById('to-pca-responsavel').value.trim();
    const tipoServico = document.getElementById('to-pca-tipo-servico').value.trim();
    const previsaoRetorno = document.getElementById('to-pca-previsao-retorno').value; // <-- CORRIGIDO: Era previsaoRetvisoRetorno

    console.log("2. Valores do formulário capturados:", { responsavel, tipoServico, previsaoRetorno }); // Log para depuração

    if (!responsavel || !tipoServico || !previsaoRetorno) {
        showAlert('Por favor, preencha o Responsável PCA, o Tipo de Serviço PCA e a Previsão de Retorno.', 'danger');
        console.log("3. Validação falhou."); // Log para depuração
        return;
    }

    const item = data.find(item => item.id === itemId);
    if (!item) {
        showAlert('Erro: Item não encontrado para enviar para PCA.', 'danger');
        console.log("4. Item não encontrado."); // Log para depuração
        return;
    }

    const updatedItem = {
        id: item.id,
        _rowIndex: item._rowIndex,
        dataEnvio: new Date().toISOString().split('T')[0], // Data de envio é sempre hoje
        previsaoRetorno: previsaoRetorno, // <-- CORRIGIDO: Era previsaoRetvisoRetorno
        responsavelPca: responsavel,
        tipoServicoPca: tipoServico,
        statusGlobal: 'emPca'
    };

    console.log("5. Chamando updateItemsInGoogleSheet com:", updatedItem); // Log para depuração
    try {
        await updateItemsInGoogleSheet(updatedItem); // Aguarda a conclusão da operação assíncrona
        closeModalToPca(); // Fechar o modal APENAS após o sucesso
        showAlert(`Item ${item.oe} enviado para PCA com sucesso!`, 'success');
        console.log("6. updateItemsInGoogleSheet bem-sucedido."); // Log para depuração
    } catch (error) {
        showAlert(`Erro ao enviar para PCA: ${error.message}`, 'danger');
        console.error('Erro ao enviar para PCA:', error);
        console.log("7. Erro durante a atualização."); // Log para depuração
    }
}

/**
 * Fecha o modal de enviar para PCA.
 */
function closeModalToPca() {
    document.getElementById('modal-to-pca').style.display = 'none';
    document.getElementById('form-to-pca').reset(); // Limpa o formulário
    console.log('closeModalToPca: Modal "Enviar para PCA" fechado.');
}

/**
 * Abre o modal para receber item do PCA.
 * @param {number} itemId - ID do item.
 */
function openFromPcaModal(itemId) {
    const item = data.find(item => item.id === itemId);
    if (!item) {
        showAlert('Erro: Item não encontrado.', 'danger');
        return;
    }
    document.getElementById('from-pca-item-id').value = item.id;
    document.getElementById('from-pca-oe').textContent = item.oe;
    document.getElementById('from-pca-data-retorno').value = formatToInputDate(new Date()); // Preenche com a data atual
    document.getElementById('from-pca-destino').value = item.destinoAposPca || ''; // Mantém o destino anterior se houver
    document.getElementById('modal-from-pca').style.display = 'flex'; // Usar 'flex' para centralização via CSS
}

/**
 * Lida com o envio do formulário de receber do PCA.
 * @param {Event} event - Evento de submit.
 */
async function handleReceiveFromPca(event) {
    event.preventDefault();
    const itemId = parseInt(document.getElementById('from-pca-item-id').value);
    const dataRetorno = document.getElementById('from-pca-data-retorno').value;
    const destino = document.getElementById('from-pca-destino').value;

    if (!dataRetorno || !destino) {
        showAlert('Por favor, preencha a Data de Retorno e o Destino Após PCA.', 'danger');
        return;
    }

    const item = data.find(item => item.id === itemId);
    if (!item) {
        showAlert('Erro: Item não encontrado para receber do PCA.', 'danger');
        return;
    }

    const updatedItem = {
        id: item.id,
        _rowIndex: item._rowIndex,
        dataRetorno: dataRetorno,
        destinoAposPca: destino,
        statusGlobal: destino === 'concluido' ? 'concluido' : destino
    };
    
    try {
        await updateItemsInGoogleSheet(updatedItem);
        closeModalFromPca(); // Fechar o modal APENAS após o sucesso
        showAlert(`Item ${item.oe} recebido do PCA com destino para "${getStatusLabel(updatedItem.statusGlobal)}"!`, 'success');
    } catch (error) {
        showAlert(`Erro ao receber do PCA: ${error.message}`, 'danger');
        console.error('Erro ao receber do PCA:', error);
    }
}

/**
 * Fecha o modal de receber do PCA.
 */
function closeModalFromPca() {
    document.getElementById('modal-from-pca').style.display = 'none';
    document.getElementById('form-from-pca').reset(); // Limpa o formulário
    console.log('closeModalFromPca: Modal "Receber do PCA" fechado.');
}