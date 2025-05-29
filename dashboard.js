const DashboardApp = {
  // --- Configuration ---
  photoMap: {
    'daniele fernanda cassaro lemes da cruz':
      'https://i.postimg.cc/NLLShyxJ/Daniele-jpg.png',
    'danielle rossettim':
      'https://i.postimg.cc/nXfPQ3fg/Danielle-jpg.png',
    'isabely tomazeli':
      'https://i.postimg.cc/F7L6R280/Isabely-jpg.png',
    'marcela cavalcante rodrigues':
      'https://i.postimg.cc/YjVVtpnp/Marcela-jpg.png',
    'marillise silvaes moraes':
      'https://i.postimg.cc/jWs3MC5h/Marillise-jpg.png',
    'milena oliveira':
      'https://i.postimg.cc/R3Xsy9Jx/Milena-jpg.png',
    'munique stein scheffel':
      'https://i.postimg.cc/JtLLvs5C/Munique-jpg.png',
  },
  nameAliasMap: {
    'danielle rossetti m.': 'danielle rossettim',
    'daniele fernanda cassaro lemes':
      'daniele fernanda cassaro lemes da cruz',
    'marillise silvaes': 'marillise silvaes moraes',
    'munique stein': 'munique stein scheffel',
  },
  statusAbertos: ['Em Andamento', 'Aguardando', 'Confirmado', 'Marcado'],
  statusAtendido: 'Atendido',
  // highlightThreshold: 10, // Removido - agora é lido do input

  // --- State ---
  allData: [],
  currentSort: 'nome_asc',
  currentSearchTerm: '',
  currentHighlightThreshold: 10, // Valor inicial, será atualizado pelo input
  detailModalInstance: null,

  // --- DOM Elements ---
  elements: {},

  // --- Initialization ---
  init() {
    this.cacheDOMElements();
    this.initializeModal();
    this.bindEvents();
    this.updateHighlightThresholdFromInput(); // Lê valor inicial do input
    this.render();
    console.log("Dashboard App Initialized");
  },

  cacheDOMElements() {
    this.elements.clinicSelect = document.getElementById('clinicSelect');
    this.elements.fileInput = document.getElementById('fileInput');
    this.elements.kpiPanel = document.getElementById('kpiPanel');
    this.elements.clinicCards = document.getElementById('clinicCards');
    this.elements.messageArea = document.getElementById('messageArea');
    this.elements.clearButton = document.getElementById('clearButton');
    this.elements.sortSelect = document.getElementById('sortSelect');
    this.elements.exportCsvButton = document.getElementById('exportCsvButton');
    this.elements.searchInput = document.getElementById('searchInput');
    this.elements.highlightThresholdInput = document.getElementById('highlightThresholdInput'); // Cache do input de limite
    // Modal elements
    this.elements.modal = document.getElementById('professionalDetailModal');
    this.elements.modalLabel = document.getElementById('professionalDetailModalLabel');
    this.elements.modalBody = document.getElementById('professionalDetailModalBody');
  },

  initializeModal() {
      if (this.elements.modal && typeof bootstrap !== 'undefined') {
          this.detailModalInstance = new bootstrap.Modal(this.elements.modal);
      } else {
          console.error("Modal element not found or Bootstrap not loaded.");
      }
  },

  bindEvents() {
    this.elements.fileInput.addEventListener('change', this.handleFileUpload.bind(this));
    if (this.elements.clearButton) {
        this.elements.clearButton.addEventListener('click', this.clearData.bind(this));
    }
    if (this.elements.sortSelect) {
        this.elements.sortSelect.addEventListener('change', this.handleSortChange.bind(this));
    }
    if (this.elements.clinicCards) {
        this.elements.clinicCards.addEventListener('click', this.handleCardClick.bind(this));
    }
    if (this.elements.exportCsvButton) {
        this.elements.exportCsvButton.addEventListener('click', this.exportSummaryToCsv.bind(this));
    }
    if (this.elements.searchInput) {
        this.elements.searchInput.addEventListener('input', this.handleSearchInput.bind(this));
    }
    if (this.elements.highlightThresholdInput) {
        this.elements.highlightThresholdInput.addEventListener('input', this.handleHighlightThresholdChange.bind(this)); // Evento para limite
    }
  },

  // --- Event Handlers ---
  handleSortChange(event) {
      this.currentSort = event.target.value;
      this.renderProfessionalCards();
  },

  handleSearchInput(event) {
      this.currentSearchTerm = event.target.value.toLowerCase().trim();
      this.renderProfessionalCards();
  },

  handleHighlightThresholdChange(event) {
      this.updateHighlightThresholdFromInput();
      this.renderProfessionalCards(); // Re-renderiza para aplicar novo limite
  },

  handleCardClick(event) {
      const cardElement = event.target.closest('.card-prof');
      if (cardElement && cardElement.dataset.profKey) {
          const professionalKey = cardElement.dataset.profKey;
          this.showProfessionalDetails(professionalKey);
      }
  },

  // --- Data Handling ---
  updateHighlightThresholdFromInput() {
      if (this.elements.highlightThresholdInput) {
          const value = parseInt(this.elements.highlightThresholdInput.value, 10);
          this.currentHighlightThreshold = isNaN(value) || value < 0 ? 0 : value; // Garante que é um número >= 0
      }
  },

  normalizeName(name) {
    if (!name || typeof name !== 'string') return '';
    const norm = name
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
    return this.nameAliasMap[norm] || norm;
  },

  handleFileUpload(event) {
    const file = event.target.files[0];
    const clinic = this.elements.clinicSelect.value;

    this.displayMessage('', '');

    if (!file || !clinic) {
      this.displayMessage('Por favor, selecione a clínica e um arquivo Excel válido.', 'danger');
      this.elements.fileInput.value = '';
      return;
    }

    this.displayMessage(`Processando arquivo "${file.name}" para a clínica ${clinic}...`, 'info');
    this.elements.fileInput.disabled = true;
    this.elements.clinicSelect.disabled = true;
    this.elements.sortSelect.disabled = true;
    this.elements.searchInput.disabled = true;
    this.elements.highlightThresholdInput.disabled = true; // Desabilita limite durante carregamento
    this.elements.exportCsvButton.disabled = true;

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            throw new Error("Nenhuma planilha encontrada no arquivo.");
        }
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        if (!sheet) {
            throw new Error(`Planilha "${workbook.SheetNames[0]}" não pôde ser lida.`);
        }

        const parsedData = XLSX.utils.sheet_to_json(sheet, { range: 1 });
        
        if (!Array.isArray(parsedData)) {
            throw new Error("Falha ao converter dados da planilha.");
        }

        const potentialDateCol = Object.keys(parsedData[0] || {}).find(k => k.toLowerCase().includes('data'));
        const potentialPatientCol = Object.keys(parsedData[0] || {}).find(k => k.toLowerCase().includes('paciente'));

        const cleanedData = parsedData.map((item) => {
          const normalized = {};
          Object.keys(item).forEach((key) => {
            const cleanKey = key ? key.toString().trim() : '';
            normalized[cleanKey] = item[key];
          });
          if (!normalized.hasOwnProperty('Profissional')) normalized['Profissional'] = '';
          if (!normalized.hasOwnProperty('Status')) normalized['Status'] = '';
          if (potentialDateCol && normalized.hasOwnProperty(potentialDateCol)) normalized['DataAtendimento'] = normalized[potentialDateCol];
          if (potentialPatientCol && normalized.hasOwnProperty(potentialPatientCol)) normalized['Paciente'] = normalized[potentialPatientCol];
          return normalized;
        });

        const newData = cleanedData.map((item) => ({
          ...item,
          Profissional: item.Profissional || '',
          Clinica: clinic,
        }));

        this.allData = this.allData.concat(newData);
        this.render();
        this.displayMessage(`Arquivo "${file.name}" processado com sucesso! ${newData.length} registros adicionados.`, 'success');

      } catch (error) {
        console.error("Erro ao processar o arquivo Excel:", error);
        this.displayMessage(`Erro ao processar o arquivo "${file.name}": ${error.message}. Verifique o formato e o conteúdo.`, 'danger');
      } finally {
        this.elements.fileInput.value = '';
        this.elements.fileInput.disabled = false;
        this.elements.clinicSelect.disabled = false;
        this.elements.sortSelect.disabled = false;
        this.elements.searchInput.disabled = false;
        this.elements.highlightThresholdInput.disabled = false; // Habilita limite
        this.elements.exportCsvButton.disabled = this.allData.length === 0;
      }
    };

    reader.onerror = () => {
      console.error("Erro ao ler o arquivo.");
      this.displayMessage(`Erro fatal ao tentar ler o arquivo "${file.name}".`, 'danger');
      this.elements.fileInput.value = '';
      this.elements.fileInput.disabled = false;
      this.elements.clinicSelect.disabled = false;
      this.elements.sortSelect.disabled = false;
      this.elements.searchInput.disabled = false;
      this.elements.highlightThresholdInput.disabled = false; // Habilita limite em caso de erro
      this.elements.exportCsvButton.disabled = true;
    };

    reader.readAsArrayBuffer(file);
  },

  clearData() {
      this.allData = [];
      this.render();
      this.displayMessage("Dados limpos. Carregue um novo arquivo.", "info");
      this.elements.clinicSelect.value = '';
      this.elements.sortSelect.value = 'nome_asc';
      this.currentSort = 'nome_asc';
      this.elements.searchInput.value = '';
      this.currentSearchTerm = '';
      this.elements.highlightThresholdInput.value = 10; // Reseta limite para o padrão
      this.updateHighlightThresholdFromInput(); // Atualiza estado interno
      // Desabilita controles
      this.elements.searchInput.disabled = true;
      this.elements.highlightThresholdInput.disabled = true;
      this.elements.exportCsvButton.disabled = true;
      this.elements.sortSelect.disabled = true;
      this.elements.clearButton.disabled = true;
  },

  // --- Data Aggregation, Filtering, and Sorting ---
  getFilteredAndSortedProfessionals() {
    const profissionais = {};
    // Agrega dados
    this.allData.forEach(item => {
      const profKey = this.normalizeName(item.Profissional);
      if (!profissionais[profKey]) {
        profissionais[profKey] = {
          key: profKey,
          nome: item.Profissional || 'Não especificado',
          total: 0,
          atendidos: 0,
          abertos: 0,
          clinicas: new Set(),
        };
      }
      profissionais[profKey].total++;
      if (item.Status === this.statusAtendido) {
          profissionais[profKey].atendidos++;
      }
      if (item.Status && this.statusAbertos.includes(item.Status)) {
          profissionais[profKey].abertos++;
      }
      if (item.Clinica) profissionais[profKey].clinicas.add(item.Clinica);
    });

    let profissionaisArray = Object.values(profissionais);

    // Filtra por busca (se houver termo)
    if (this.currentSearchTerm) {
        profissionaisArray = profissionaisArray.filter(prof => 
            prof.nome.toLowerCase().includes(this.currentSearchTerm)
        );
    }

    // Ordena o array filtrado
    profissionaisArray.sort((a, b) => {
      const percentA = a.total > 0 ? (a.atendidos / a.total) : 0;
      const percentB = b.total > 0 ? (b.atendidos / b.total) : 0;
      const naoFinalizadosA = a.total - a.atendidos;
      const naoFinalizadosB = b.total - b.atendidos;

      switch (this.currentSort) {
        case 'nome_asc': return a.nome.localeCompare(b.nome);
        case 'nome_desc': return b.nome.localeCompare(a.nome);
        case 'total_desc': return b.total - a.total;
        case 'total_asc': return a.total - b.total;
        case 'finalizados_desc': return percentB - percentA;
        case 'finalizados_asc': return percentA - percentB;
        case 'nao_finalizados_desc': return naoFinalizadosB - naoFinalizadosA;
        case 'nao_finalizados_asc': return naoFinalizadosA - naoFinalizadosB;
        case 'abertos_desc': return b.abertos - a.abertos;
        case 'abertos_asc': return a.abertos - b.abertos;
        default: return a.nome.localeCompare(b.nome);
      }
    });

    return profissionaisArray;
  },

  // --- CSV Export ---
  exportSummaryToCsv() {
      const professionalsToExport = this.getFilteredAndSortedProfessionals(); 

      if (professionalsToExport.length === 0) {
          this.displayMessage("Não há dados (filtrados) para exportar.", "warning");
          return;
      }

      let csvContent = "data:text/csv;charset=utf-8,\uFEFF";

      csvContent += "Profissional;Clinicas;Total Atendimentos;Finalizados;Nao Finalizados;Abertos;% Finalizados\r\n";

      professionalsToExport.forEach(prof => {
          const clinicasStr = [...prof.clinicas].join(', ');
          const percent = prof.total > 0 ? Math.round((prof.atendidos / prof.total) * 100) : 0;
          const naoFinalizados = prof.total - prof.atendidos;
          
          const escapeCsvField = (field) => {
              const fieldStr = String(field).replace(/"/g, '""');
              return fieldStr.includes(';') || fieldStr.includes('\"') || fieldStr.includes('\r') || fieldStr.includes('\n') ? `"${fieldStr}"` : fieldStr;
          };

          const row = [
              escapeCsvField(prof.nome),
              escapeCsvField(clinicasStr),
              prof.total,
              prof.atendidos,
              naoFinalizados,
              prof.abertos,
              `${percent}%`
          ].join(";");
          csvContent += row + "\r\n";
      });

      const encodedUri = encodeURI(csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", "resumo_atendimentos_profissionais.csv");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      this.displayMessage("Resumo exportado para CSV.", "success");
  },

  // --- UI Rendering ---
  displayMessage(message, type) {
    if (!this.elements.messageArea) return;
    
    this.elements.messageArea.innerHTML = '';
    if (!message) return;

    const alertDiv = document.createElement('div');
    alertDiv.className = `alert-custom alert-${type}-custom`;
    alertDiv.textContent = message;
    alertDiv.setAttribute('role', 'alert');

    this.elements.messageArea.appendChild(alertDiv);

    if (type === 'success' || type === 'info') {
        setTimeout(() => {
            if (alertDiv.parentNode === this.elements.messageArea) {
                 this.elements.messageArea.innerHTML = '';
            }
        }, 5000);
    }
  },

  render() {
    this.renderKpis();
    this.renderProfessionalCards();
    // Habilita/desabilita controles após renderizar
    const hasData = this.allData.length > 0;
    if (this.elements.exportCsvButton) this.elements.exportCsvButton.disabled = !hasData;
    if (this.elements.searchInput) this.elements.searchInput.disabled = !hasData;
    if (this.elements.sortSelect) this.elements.sortSelect.disabled = !hasData;
    if (this.elements.clearButton) this.elements.clearButton.disabled = !hasData;
    if (this.elements.highlightThresholdInput) this.elements.highlightThresholdInput.disabled = !hasData; // Habilita/desabilita limite
  },

  renderKpis() {
    const panel = this.elements.kpiPanel;
    panel.innerHTML = '';

    if (this.allData.length === 0) {
      panel.innerHTML = '<p class="text-center text-muted">Nenhum dado carregado ainda. Selecione uma clínica e anexe uma planilha.</p>';
      return;
    }

    const data = this.allData;
    const totalAbertosGeral = data.filter(d => d.Status && this.statusAbertos.includes(d.Status)).length;

    const porProfAbertos = {};
    const porClinicaAbertos = {};
    data.forEach(d => {
      if (d.Status && this.statusAbertos.includes(d.Status)) {
        const prof = d.Profissional || 'Não especificado';
        const clinica = d.Clinica || 'Não especificada';
        porProfAbertos[prof] = (porProfAbertos[prof] || 0) + 1;
        porClinicaAbertos[clinica] = (porClinicaAbertos[clinica] || 0) + 1;
      }
    });

    const top3Prof = Object.entries(porProfAbertos).sort((a, b) => b[1] - a[1]).slice(0, 3);
    const top3Clin = Object.entries(porClinicaAbertos).sort((a, b) => b[1] - a[1]).slice(0, 3);

    const fragment = document.createDocumentFragment();

    const totalCard = document.createElement('div');
    totalCard.className = 'card';
    totalCard.innerHTML = `<h5>Total Atendimentos Abertos</h5><p>${totalAbertosGeral}</p>`;
    fragment.appendChild(totalCard);

    const profCard = document.createElement('div');
    profCard.className = 'card';
    let profListHTML = top3Prof.map(([prof, qt]) => `<li>${prof} (${qt})</li>`).join('');
    profCard.innerHTML = `<h5>Top 3 Profissionais (Abertos)</h5><ul>${profListHTML || '<li>N/A</li>'}</ul>`;
    fragment.appendChild(profCard);

    const clinicCard = document.createElement('div');
    clinicCard.className = 'card';
    let clinicListHTML = top3Clin.map(([clin, qt]) => `<li>${clin} (${qt})</li>`).join('');
    clinicCard.innerHTML = `<h5>Top 3 Clínicas (Abertos)</h5><ul>${clinicListHTML || '<li>N/A</li>'}</ul>`;
    fragment.appendChild(clinicCard);

    panel.appendChild(fragment);
  },

  renderProfessionalCards() {
    const container = this.elements.clinicCards;
    container.innerHTML = '';

    const profissionaisFiltradosOrdenados = this.getFilteredAndSortedProfessionals();

    if (profissionaisFiltradosOrdenados.length === 0) {
        if (this.allData.length > 0) {
             container.innerHTML = '<p class="text-center text-muted">Nenhum profissional encontrado para o termo buscado.</p>';
        }
      return;
    }

    const fragment = document.createDocumentFragment();

    profissionaisFiltradosOrdenados.forEach(info => {
      const img = this.photoMap[info.key] || 'https://via.placeholder.com/120x120?text=Sem+Foto';
      const percent = info.total > 0 ? Math.round((info.atendidos / info.total) * 100) : 0;
      const naoFinalizados = info.total - info.atendidos;

      const card = document.createElement('div');
      card.className = 'card-prof';
      card.dataset.profKey = info.key;
      card.style.cursor = 'pointer';

      // Usa o valor do estado atualizado pelo input
      if (info.abertos > this.currentHighlightThreshold) { 
          card.classList.add('highlight-attention');
      } else {
          card.classList.remove('highlight-attention'); // Garante que remove se não atender mais
      }

      card.innerHTML = `
        <img src="${img}" alt="Foto de ${info.nome}" loading="lazy" />
        <div class="prof-name">${info.nome}</div>
        <div class="clinic-name">${info.clinicas.size > 0 ? [...info.clinicas].join(', ') : 'N/A'}</div>
        <div class="count">Total: ${info.total}</div>
        <div class="count text-success">Finalizados: ${info.atendidos}</div>
        <div class="count text-warning">Não Finalizados: ${naoFinalizados}</div> 
        <div class="progress mt-2" aria-label="Progresso de atendimentos finalizados">
          <div class="progress-bar" style="width: ${percent}%" role="progressbar" aria-valuenow="${percent}" aria-valuemin="0" aria-valuemax="100"></div>
        </div>
        <small>${percent}% finalizados</small>
      `;
      fragment.appendChild(card);
    });

    container.appendChild(fragment);
  },

  showProfessionalDetails(professionalKey) {
      if (!this.detailModalInstance) {
          console.error("Modal instance not available.");
          return;
      }

      const allProfessionalsAggregated = this.getFilteredAndSortedProfessionals(); 
      let professionalData = allProfessionalsAggregated.find(p => p.key === professionalKey);
      
      if (!professionalData) {
          const allAggregated = this.aggregateAllProfessionals(); 
          professionalData = allAggregated.find(p => p.key === professionalKey);
      }

      if (!professionalData) {
          console.error("Professional data not found for key:", professionalKey);
          const originalRecord = this.allData.find(item => this.normalizeName(item.Profissional) === professionalKey);
          if (originalRecord) {
              professionalData = { nome: originalRecord.Profissional };
          } else {
              this.elements.modalLabel.textContent = `Detalhes de Atendimentos`;
              this.elements.modalBody.innerHTML = '<p>Erro: Dados do profissional não encontrados.</p>';
              this.detailModalInstance.show();
              return;
          }
      }

      const appointments = this.allData.filter(item => this.normalizeName(item.Profissional) === professionalKey);

      this.elements.modalLabel.textContent = `Detalhes de Atendimentos - ${professionalData.nome}`;

      let tableHTML = '<p>Nenhum atendimento encontrado.</p>';
      if (appointments.length > 0) {
          const hasDate = appointments[0].hasOwnProperty('DataAtendimento');
          const hasPatient = appointments[0].hasOwnProperty('Paciente');

          tableHTML = `
              <table class="table table-striped table-hover table-sm">
                  <thead>
                      <tr>
                          <th>Status</th>
                          <th>Clínica</th>
                          ${hasDate ? '<th>Data</th>' : ''}
                          ${hasPatient ? '<th>Paciente</th>' : ''}
                      </tr>
                  </thead>
                  <tbody>
          `;
          appointments.forEach(appt => {
              let formattedDate = '';
              if (hasDate && appt.DataAtendimento) {
                  if (typeof appt.DataAtendimento === 'number') {
                      try {
                          const excelEpoch = new Date(1899, 11, 30);
                          const jsDate = new Date(excelEpoch.getTime() + (appt.DataAtendimento - 1) * 24 * 60 * 60 * 1000);
                          formattedDate = jsDate.toLocaleDateString('pt-BR');
                      } catch (e) { formattedDate = 'Data inválida'; }
                  } else {
                      try {
                          formattedDate = new Date(appt.DataAtendimento).toLocaleDateString('pt-BR');
                      } catch (e) { formattedDate = appt.DataAtendimento; }
                  }
              }
              
              tableHTML += `
                  <tr>
                      <td>${appt.Status || 'N/A'}</td>
                      <td>${appt.Clinica || 'N/A'}</td>
                      ${hasDate ? `<td>${formattedDate || 'N/A'}</td>` : ''}
                      ${hasPatient ? `<td>${appt.Paciente || 'N/A'}</td>` : ''}
                  </tr>
              `;
          });
          tableHTML += `
                  </tbody>
              </table>
          `;
      }

      this.elements.modalBody.innerHTML = tableHTML;
      this.detailModalInstance.show();
  },

  // Helper para agregar todos (usado no fallback do modal)
  aggregateAllProfessionals() { 
     const profissionais = {};
     this.allData.forEach(item => {
       const profKey = this.normalizeName(item.Profissional);
       if (!profissionais[profKey]) {
         profissionais[profKey] = {
           key: profKey,
           nome: item.Profissional || 'Não especificado',
           total: 0,
           atendidos: 0,
           abertos: 0,
           clinicas: new Set(),
         };
       }
       profissionais[profKey].total++;
       if (item.Status === this.statusAtendido) {
           profissionais[profKey].atendidos++;
       }
       if (item.Status && this.statusAbertos.includes(item.Status)) {
           profissionais[profKey].abertos++;
       }
       if (item.Clinica) profissionais[profKey].clinicas.add(item.Clinica);
     });
     return Object.values(profissionais);
  }
};

// --- Initialize the app once the DOM is ready ---
document.addEventListener('DOMContentLoaded', () => {
  if (typeof bootstrap !== 'undefined') {
      DashboardApp.init();
  } else {
      console.error("Bootstrap JS not found. Modal functionality might be affected.");
      DashboardApp.init(); 
  }
});

