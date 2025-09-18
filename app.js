/**
 * Personnel Organizational Chart - With Data Sources
 * Auto-loads personnel_data.csv on first visit
 * CSV Format: name,email,role,team,institution,projects,data_sources,services,status
 */

// ================================
// GLOBAL STATE
// ================================
const AppState = {
    rawData: [],           // Original CSV data
    filteredData: [],      // Filtered data for display
    teams: new Set(),      // Unique teams
    institutions: new Set(), // Unique institutions
    projects: new Set(),    // Unique projects
    dataSources: new Set(), // Unique data sources
    services: new Set(),    // Unique services
    currentView: 'grid',    // Current view mode
    currentPage: 1,         // Current page for table view
    itemsPerPage: 10,       // Items per page
    sortColumn: null,       // Current sort column
    sortDirection: 'asc',   // Sort direction
    csvHeaders: [],         // CSV column headers
    hasData: false,         // Whether data is loaded
    csvFileName: 'personnel_data.csv'  // Default CSV file to load
};

// ================================
// INITIALIZATION
// ================================
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

async function initializeApp() {
    // Initialize event listeners
    initializeEventListeners();
    
    // Show loading spinner
    showLoadingSpinner(true);
    
    // Try to load saved data from localStorage
    const hasData = loadSavedData();
    
    if (hasData) {
        // Data exists in localStorage - show dashboard
        showDashboard();
        updateStatistics();
        renderCurrentView();
        showLoadingSpinner(false);
        
        // Show last update info
        const lastUpdate = localStorage.getItem('orgChartLastUpdate');
        if (lastUpdate) {
            const date = new Date(lastUpdate);
            const timeAgo = getTimeAgo(date);
            updateDataInfoBar(`Data loaded: ${AppState.rawData.length} personnel • Last updated: ${timeAgo}`);
        }
    } else {
        // No saved data - try to auto-load CSV file
        try {
            await loadCSVFromFile();
        } catch (error) {
            console.log('Could not auto-load CSV file, loading sample data instead');
            loadSampleData();
        }
        showLoadingSpinner(false);
    }
}

// ================================
// AUTO-LOAD CSV FILE
// ================================
async function loadCSVFromFile() {
    try {
        // Try to fetch the CSV file from the same directory
        const response = await fetch(AppState.csvFileName);
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const csvText = await response.text();
        
        // Check if we got actual CSV content
        if (!csvText || csvText.trim().length === 0) {
            throw new Error('Empty CSV file');
        }
        
        // Parse the CSV
        const csvData = parseCSV(csvText);
        
        // Validate the CSV data
        if (validateCSVData(csvData)) {
            loadCSVData(csvData);
            showDashboard();
            updateStatistics();
            renderCurrentView();
            showToast(`Loaded ${csvData.length} personnel from ${AppState.csvFileName}`, 'success');
            updateDataInfoBar(`Data loaded: ${csvData.length} personnel from ${AppState.csvFileName}`);
        }
    } catch (error) {
        console.error('Error loading CSV file:', error);
        
        // If we're using the file:// protocol, show a helpful message
        if (window.location.protocol === 'file:') {
            showToast('Direct file access not allowed. Please run this on a web server or use the Upload button.', 'warning');
        }
        
        throw error; // Re-throw to be caught by the calling function
    }
}

// ================================
// SAMPLE DATA
// ================================
function loadSampleData() {
    const sampleCSV = `name,email,role,team,institution,projects,data_sources,services,status
Linda Chaba,person1@gmail.com,Translational TB Team Leader,Translational TB;DSMG,UCSF,Linda Project 1, Linda Project 2,JHU;CSU,Slack;Teams,active
Ziran Li,person2@gmail.com,Malaria Team Leader,Malaria; Translational TB,UCSF,Ziran Project 1, Ziran Project 2,JHU;CSU; Servers,GitHub;Docker,active
Person3,person3@company.com,Senior Researcher,Research;Data Science,Harvard University,Machine Learning;Data Pipeline,Research Data;Survey Results;Lab Results,Python;Jupyter;AWS,active
Person4,person4@company.com,Developer,Engineering;DevOps,Stanford University,API Platform;Mobile App,Application Logs;User Analytics,GitHub;Jenkins;Docker,active
Person5,person5@company.com,Junior Developer,Engineering,MIT,Mobile App;Web Portal,Test Data;Development Database,GitHub;VS Code,active
Person6,person6@company.com,Project Manager,Management;Operations,Yale University,Mobile App;Customer Dashboard,Project Metrics;Budget Reports,Jira;Slack;Confluence,active
Person7,person7@company.com,Data Analyst,Data Science;Analytics,Harvard University,Data Pipeline;Analytics Report,Customer Data;Sales Data;Marketing Analytics,Tableau;SQL Server;Excel,active
Person8,person8@company.com,DevOps Engineer,DevOps;Infrastructure,Stanford University,Cloud Migration;Infrastructure Upgrade,Infrastructure Logs;Security Logs;Deployment Metrics,Docker;Jenkins;Kubernetes;AWS,active`;
    
    const csvData = parseCSV(sampleCSV);
    loadCSVData(csvData);
    showDashboard();
    updateStatistics();
    renderCurrentView();
    showToast('Sample data loaded (place personnel_data.csv in same directory for auto-load)', 'info');
    updateDataInfoBar(`Sample data loaded: ${csvData.length} personnel`);
}

// ================================
// MANUAL CSV REFRESH
// ================================
async function refreshData() {
    if (AppState.hasData) {
        // Try to reload the CSV file
        showLoadingSpinner(true);
        try {
            await loadCSVFromFile();
            showToast('Data refreshed from CSV file', 'success');
        } catch (error) {
            // If CSV load fails, just refresh the existing data
            extractUniqueValues(AppState.rawData);
            populateFilters();
            applyFilters();
            updateStatistics();
            renderCurrentView();
            showToast('Data refreshed', 'success');
        }
        showLoadingSpinner(false);
    } else {
        showToast('No data to refresh', 'warning');
    }
}

// ================================
// EVENT LISTENERS
// ================================
function initializeEventListeners() {
    // CSV Upload (Admin only)
    const uploadArea = document.getElementById('csvUploadArea');
    const fileInput = document.getElementById('csvFileInput');
    const hiddenInput = document.getElementById('hiddenFileInput');
    
    if (uploadArea) {
        uploadArea.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('dragover', handleDragOver);
        uploadArea.addEventListener('dragleave', handleDragLeave);
        uploadArea.addEventListener('drop', handleDrop);
    }
    
    if (fileInput) {
        fileInput.addEventListener('change', handleFileSelect);
    }
    
    if (hiddenInput) {
        hiddenInput.addEventListener('change', handleFileSelect);
    }
    
    // Navigation buttons
    document.getElementById('refreshData').addEventListener('click', refreshData);
    
    const uploadBtn = document.getElementById('uploadCSV');
    if (uploadBtn) {
        uploadBtn.addEventListener('click', showImportSection);
    }
    
    const cancelImportBtn = document.getElementById('cancelImport');
    if (cancelImportBtn) {
        cancelImportBtn.addEventListener('click', hideImportSection);
    }
    
    // View tabs
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const target = e.target.closest('.tab-btn');
            if (target) {
                switchView(target.dataset.view);
            }
        });
    });
    
    // Search and filters
    document.getElementById('searchInput').addEventListener('input', debounce(applyFilters, 300));
    document.getElementById('teamFilter').addEventListener('change', applyFilters);
    document.getElementById('institutionFilter').addEventListener('change', applyFilters);
    document.getElementById('projectFilter').addEventListener('change', applyFilters);
    document.getElementById('dataSourceFilter').addEventListener('change', applyFilters);
    document.getElementById('statusFilter').addEventListener('change', applyFilters);
    document.getElementById('clearFilters').addEventListener('click', clearFilters);
    
    // Export buttons
    document.getElementById('exportCSV').addEventListener('click', exportAsCSV);
    document.getElementById('exportJSON').addEventListener('click', exportAsJSON);
    document.getElementById('exportPDF').addEventListener('click', exportAsPDF);
    document.getElementById('printReport').addEventListener('click', printReport);
    
    // Modal controls
    document.getElementById('closeModal').addEventListener('click', closePersonModal);
    document.getElementById('closeModalBtn').addEventListener('click', closePersonModal);
    document.getElementById('closeProjectModal').addEventListener('click', closeProjectModal);
    document.getElementById('closeProjectModalBtn').addEventListener('click', closeProjectModal);
    document.getElementById('closeDataSourceModal').addEventListener('click', closeDataSourceModal);
    document.getElementById('closeDataSourceModalBtn').addEventListener('click', closeDataSourceModal);
}

// ================================
// DATA LOADING
// ================================
function loadSavedData() {
    const saved = localStorage.getItem('orgChartData');
    if (saved) {
        try {
            const data = JSON.parse(saved);
            if (data.rawData && data.rawData.length > 0) {
                AppState.csvHeaders = data.csvHeaders || [];
                AppState.rawData = data.rawData;
                AppState.filteredData = data.rawData;
                AppState.hasData = true;
                
                // Extract unique values for filters
                extractUniqueValues(data.rawData);
                
                // Populate filters
                populateFilters();
                
                return true;
            }
        } catch (error) {
            console.error('Error loading saved data:', error);
        }
    }
    return false;
}

function loadCSVData(data) {
    AppState.rawData = data;
    AppState.filteredData = data;
    AppState.hasData = true;
    
    // Extract unique values for filters
    extractUniqueValues(data);
    
    // Populate filters
    populateFilters();
    
    // Save to localStorage
    saveToLocalStorage();
    
    // Update display
    updateStatistics();
    renderCurrentView();
}

// ================================
// CSV FILE HANDLING
// ================================
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.classList.add('drag-over');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.classList.remove('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.classList.remove('drag-over');
    
    const files = e.dataTransfer.files;
    if (files.length > 0 && files[0].type === 'text/csv') {
        processCSVFile(files[0]);
    } else {
        showToast('Please drop a valid CSV file', 'error');
    }
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file && file.type === 'text/csv') {
        processCSVFile(file);
    } else if (file) {
        showToast('Please select a valid CSV file', 'error');
    }
}

function processCSVFile(file) {
    showLoadingSpinner(true);
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const csvData = parseCSV(e.target.result);
            if (validateCSVData(csvData)) {
                loadCSVData(csvData);
                showToast('Data updated successfully', 'success');
                updateDataInfoBar(`Data loaded: ${csvData.length} personnel from ${file.name}`);
                hideImportSection();
            }
        } catch (error) {
            showToast('Error processing CSV file: ' + error.message, 'error');
        } finally {
            showLoadingSpinner(false);
        }
    };
    reader.readAsText(file);
}

// ================================
// CSV PARSING
// ================================
function parseCSV(csvText) {
    const lines = csvText.split('\n').filter(line => line.trim());
    if (lines.length < 2) {
        throw new Error('CSV file must contain headers and at least one data row');
    }
    
    const headers = parseCSVLine(lines[0]);
    AppState.csvHeaders = headers;
    
    const data = [];
    for (let i = 1; i < lines.length; i++) {
        const values = parseCSVLine(lines[i]);
        const row = {};
        headers.forEach((header, index) => {
            row[header.toLowerCase().replace(/\s+/g, '_')] = values[index] || '';
        });
        data.push(row);
    }
    
    return data;
}

function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        const nextChar = line[i + 1];
        
        if (char === '"') {
            if (inQuotes && nextChar === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === ',' && !inQuotes) {
            result.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    
    result.push(current.trim());
    return result;
}

function validateCSVData(data) {
    if (!data || data.length === 0) {
        throw new Error('No data found in CSV file');
    }
    
    // Check for required columns
    const requiredColumns = ['name', 'email', 'role', 'team'];
    const headers = Object.keys(data[0]);
    
    for (const column of requiredColumns) {
        if (!headers.includes(column)) {
            throw new Error(`Required column "${column}" not found in CSV`);
        }
    }
    
    // Validate each row has required data
    data.forEach((row, index) => {
        if (!row.name || !row.email || !row.role || !row.team) {
            throw new Error(`Row ${index + 2} is missing required data`);
        }
    });
    
    return true;
}

// ================================
// DATA EXTRACTION
// ================================
function extractUniqueValues(data) {
    AppState.teams.clear();
    AppState.institutions.clear();
    AppState.projects.clear();
    AppState.dataSources.clear();
    AppState.services.clear();
    
    data.forEach(person => {
        // Teams (semicolon-separated)
        if (person.team) {
            person.team.split(';').forEach(team => {
                AppState.teams.add(team.trim());
            });
        }
        
        // Institutions
        if (person.institution) {
            AppState.institutions.add(person.institution.trim());
        }
        
        // Projects (semicolon-separated)
        if (person.projects) {
            person.projects.split(';').forEach(project => {
                AppState.projects.add(project.trim());
            });
        }
        
        // Data Sources (semicolon-separated)
        if (person.data_sources) {
            person.data_sources.split(';').forEach(source => {
                AppState.dataSources.add(source.trim());
            });
        }
        
        // Services (semicolon-separated)
        if (person.services) {
            person.services.split(';').forEach(service => {
                AppState.services.add(service.trim());
            });
        }
    });
}

// ================================
// UI UPDATES
// ================================
function showDashboard() {
    document.getElementById('mainDashboard').style.display = 'block';
    const importSection = document.getElementById('dataImportSection');
    if (importSection) {
        importSection.style.display = 'none';
    }
}

function showImportSection() {
    document.getElementById('dataImportSection').style.display = 'block';
    document.getElementById('mainDashboard').style.display = 'none';
}

function hideImportSection() {
    document.getElementById('dataImportSection').style.display = 'none';
    document.getElementById('mainDashboard').style.display = 'block';
}

function updateDataInfoBar(message) {
    const dataInfoText = document.getElementById('dataInfoText');
    if (dataInfoText) {
        dataInfoText.textContent = message;
    }
}

function updateStatistics() {
    // Basic counts
    document.getElementById('totalPersonnel').textContent = AppState.rawData.length;
    document.getElementById('totalTeams').textContent = AppState.teams.size;
    document.getElementById('totalInstitutions').textContent = AppState.institutions.size;
    document.getElementById('totalProjects').textContent = AppState.projects.size;
    document.getElementById('totalDataSources').textContent = AppState.dataSources.size;
    document.getElementById('totalServices').textContent = AppState.services.size;
}

// ================================
// VIEW MANAGEMENT
// ================================
function switchView(view) {
    AppState.currentView = view;
    
    // Update tab active states
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.view === view);
    });
    
    // Update view content
    document.querySelectorAll('.view-content').forEach(content => {
        content.classList.remove('active');
    });
    
    // Map view names to element IDs
    const viewMap = {
        'data-sources': 'dataSourcesView'
    };
    
    const viewElementId = viewMap[view] || `${view}View`;
    const viewElement = document.getElementById(viewElementId);
    
    if (viewElement) {
        viewElement.classList.add('active');
    }
    
    renderCurrentView();
}

function renderCurrentView() {
    if (!AppState.hasData) {
        return;
    }
    
    switch (AppState.currentView) {
        case 'grid':
            renderGridView();
            break;
        case 'table':
            renderTableView();
            break;
        case 'teams':
            renderTeamsView();
            break;
        case 'projects':
            renderProjectsView();
            break;
        case 'data-sources':
            renderDataSourcesView();
            break;
        case 'services':
            renderServicesView();
            break;
    }
}

// ================================
// GRID VIEW
// ================================
function renderGridView() {
    const container = document.getElementById('personnelGrid');
    const data = AppState.filteredData;
    
    if (data.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: var(--gray-500);">No personnel data to display</p>';
        return;
    }
    
    container.innerHTML = data.map(person => {
        const initials = getInitials(person.name);
        const teams = person.team ? person.team.split(';').map(t => t.trim()) : [];
        const projects = person.projects ? person.projects.split(';').map(p => p.trim()) : [];
        const dataSources = person.data_sources ? person.data_sources.split(';').map(d => d.trim()) : [];
        const status = person.status || 'active';
        
        return `
            <div class="person-card" onclick="showPersonDetails(${JSON.stringify(person).replace(/"/g, '&quot;')})">
                <div class="person-header">
                    <div class="person-avatar">${initials}</div>
                    <div class="person-info">
                        <h3>${person.name}</h3>
                        <p>${person.role}</p>
                    </div>
                </div>
                <div class="person-details">
                    <div class="detail-row">
                        <span class="detail-label">Email:</span>
                        <span class="detail-value">${person.email}</span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Team:</span>
                        <span class="detail-value">
                            ${teams.map(t => `<span class="tag tag-team">${t}</span>`).join('')}
                        </span>
                    </div>
                    ${person.institution ? `
                        <div class="detail-row">
                            <span class="detail-label">Institution:</span>
                            <span class="detail-value">${person.institution}</span>
                        </div>
                    ` : ''}
                    <div class="detail-row">
                        <span class="detail-label">Projects:</span>
                        <span class="detail-value">
                            ${projects.length > 0 ? projects.slice(0, 2).map(p => `<span class="tag tag-project">${p}</span>`).join('') : '<span style="color: var(--gray-400);">None</span>'}
                            ${projects.length > 2 ? `<span class="tag tag-more">+${projects.length - 2} more</span>` : ''}
                        </span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Data Sources:</span>
                        <span class="detail-value">
                            ${dataSources.length > 0 ? dataSources.slice(0, 2).map(d => `<span class="tag tag-data">${d}</span>`).join('') : '<span style="color: var(--gray-400);">None</span>'}
                            ${dataSources.length > 2 ? `<span class="tag tag-more">+${dataSources.length - 2} more</span>` : ''}
                        </span>
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Status:</span>
                        <span class="detail-value">
                            <span class="status-badge ${status}">${status}</span>
                        </span>
                    </div>
                </div>
            </div>
        `;
    }).join('');
}

// ================================
// TABLE VIEW
// ================================
function renderTableView() {
    const tbody = document.getElementById('personnelTableBody');
    const data = paginateData(AppState.filteredData);
    
    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" style="text-align: center; color: var(--gray-500);">No personnel data to display</td></tr>';
        renderPagination();
        return;
    }
    
    tbody.innerHTML = data.map(person => {
        const teams = person.team ? person.team.split(';').map(t => `<span class="tag tag-team">${t.trim()}</span>`).join('') : '-';
        const projects = person.projects ? person.projects.split(';').map(p => `<span class="tag tag-project">${p.trim()}</span>`).join('') : '-';
        const dataSources = person.data_sources ? person.data_sources.split(';').map(d => `<span class="tag tag-data">${d.trim()}</span>`).join('') : '-';
        const services = person.services ? person.services.split(';').map(s => `<span class="tag tag-service">${s.trim()}</span>`).join('') : '-';
        const status = person.status || 'active';
        
        return `
            <tr>
                <td><strong>${person.name}</strong></td>
                <td>${person.email}</td>
                <td>${person.role}</td>
                <td>${teams}</td>
                <td>${person.institution || '-'}</td>
                <td>${projects}</td>
                <td>${dataSources}</td>
                <td>${services}</td>
                <td><span class="status-badge ${status}">${status}</span></td>
                <td>
                    <button class="btn-icon" onclick='showPersonDetails(${JSON.stringify(person).replace(/'/g, "&apos;")})' title="View Details">
                        <i class="fas fa-eye"></i>
                    </button>
                </td>
            </tr>
        `;
    }).join('');
    
    renderPagination();
    
    // Add sort functionality to headers
    document.querySelectorAll('.data-table th[data-sort]').forEach(th => {
        th.style.cursor = 'pointer';
        th.onclick = () => sortTable(th.dataset.sort);
    });
}

function paginateData(data) {
    const start = (AppState.currentPage - 1) * AppState.itemsPerPage;
    const end = start + AppState.itemsPerPage;
    return data.slice(start, end);
}

function renderPagination() {
    const container = document.getElementById('pagination');
    const totalPages = Math.ceil(AppState.filteredData.length / AppState.itemsPerPage);
    
    if (totalPages <= 1) {
        container.innerHTML = '';
        return;
    }
    
    let html = `
        <button class="page-btn" onclick="changePage(${AppState.currentPage - 1})" 
                ${AppState.currentPage === 1 ? 'disabled' : ''}>
            <i class="fas fa-chevron-left"></i>
        </button>
    `;
    
    // Smart pagination
    const maxButtons = 5;
    let startPage = Math.max(1, AppState.currentPage - 2);
    let endPage = Math.min(totalPages, startPage + maxButtons - 1);
    
    if (endPage - startPage < maxButtons - 1) {
        startPage = Math.max(1, endPage - maxButtons + 1);
    }
    
    if (startPage > 1) {
        html += `<button class="page-btn" onclick="changePage(1)">1</button>`;
        if (startPage > 2) html += `<span>...</span>`;
    }
    
    for (let i = startPage; i <= endPage; i++) {
        html += `
            <button class="page-btn ${i === AppState.currentPage ? 'active' : ''}" 
                    onclick="changePage(${i})">${i}</button>
        `;
    }
    
    if (endPage < totalPages) {
        if (endPage < totalPages - 1) html += `<span>...</span>`;
        html += `<button class="page-btn" onclick="changePage(${totalPages})">${totalPages}</button>`;
    }
    
    html += `
        <button class="page-btn" onclick="changePage(${AppState.currentPage + 1})" 
                ${AppState.currentPage === totalPages ? 'disabled' : ''}>
            <i class="fas fa-chevron-right"></i>
        </button>
    `;
    
    container.innerHTML = html;
}

function changePage(page) {
    const totalPages = Math.ceil(AppState.filteredData.length / AppState.itemsPerPage);
    if (page < 1 || page > totalPages) return;
    
    AppState.currentPage = page;
    renderTableView();
}

// ================================
// TEAMS VIEW
// ================================
function renderTeamsView() {
    const container = document.getElementById('teamsGrid');
    
    if (AppState.teams.size === 0) {
        container.innerHTML = '<p style="grid-column: 1/-1; text-align: center; color: var(--gray-500);">No teams found</p>';
        return;
    }
    
    const teamsData = [];
    AppState.teams.forEach(teamName => {
        const members = AppState.filteredData.filter(person => 
            person.team && person.team.split(';').map(t => t.trim()).includes(teamName)
        );
        teamsData.push({ name: teamName, members });
    });
    
    // Sort teams by member count
    teamsData.sort((a, b) => b.members.length - a.members.length);
    
    container.innerHTML = teamsData.map(team => `
        <div class="team-card">
            <div class="card-header">
                <div class="card-icon">
                    <i class="fas fa-users"></i>
                </div>
                <div>
                    <div class="card-title">${team.name}</div>
                    <div class="card-subtitle">${team.members.length} members</div>
                </div>
            </div>
            <div class="member-list">
                ${team.members.slice(0, 5).map(m => 
                    `<span class="member-chip" onclick='showPersonDetails(${JSON.stringify(m).replace(/'/g, "&apos;")})'>${m.name}</span>`
                ).join('')}
                ${team.members.length > 5 ? `<span class="member-chip">+${team.members.length - 5} more</span>` : ''}
            </div>
        </div>
    `).join('');
}

// ================================
// PROJECTS VIEW
// ================================
function renderProjectsView() {
    const container = document.getElementById('projectsGrid');
    
    if (AppState.projects.size === 0) {
        container.innerHTML = '<p style="grid-column: 1/-1; text-align: center; color: var(--gray-500);">No projects found</p>';
        return;
    }
    
    const projectsData = [];
    AppState.projects.forEach(projectName => {
        const members = AppState.filteredData.filter(person => 
            person.projects && person.projects.split(';').map(p => p.trim()).includes(projectName)
        );
        
        // Count by institution
        const institutionCount = {};
        members.forEach(m => {
            const inst = m.institution || 'Unknown';
            institutionCount[inst] = (institutionCount[inst] || 0) + 1;
        });
        
        projectsData.push({ name: projectName, members, institutionCount });
    });
    
    // Sort projects by member count
    projectsData.sort((a, b) => b.members.length - a.members.length);
    
    container.innerHTML = projectsData.map(project => `
        <div class="project-card" onclick='showProjectDetails(${JSON.stringify(project).replace(/'/g, "&apos;")})'>
            <div class="card-header">
                <div class="card-icon">
                    <i class="fas fa-project-diagram"></i>
                </div>
                <div>
                    <div class="card-title">${project.name}</div>
                    <div class="card-subtitle">${project.members.length} team members</div>
                </div>
            </div>
            <div class="project-institutions">
                ${Object.entries(project.institutionCount).map(([inst, count]) => 
                    `<span class="inst-badge">${inst} (${count})</span>`
                ).join('')}
            </div>
            <div class="member-list">
                ${project.members.slice(0, 4).map(m => 
                    `<span class="member-chip">${m.name}</span>`
                ).join('')}
                ${project.members.length > 4 ? `<span class="member-chip">+${project.members.length - 4} more</span>` : ''}
            </div>
        </div>
    `).join('');
}

// ================================
// DATA SOURCES VIEW
// ================================
function renderDataSourcesView() {
    const container = document.getElementById('dataSourcesGrid');
    
    if (AppState.dataSources.size === 0) {
        container.innerHTML = '<p style="grid-column: 1/-1; text-align: center; color: var(--gray-500);">No data sources found</p>';
        return;
    }
    
    const dataSourcesData = [];
    AppState.dataSources.forEach(sourceName => {
        const owners = AppState.filteredData.filter(person => 
            person.data_sources && person.data_sources.split(';').map(d => d.trim()).includes(sourceName)
        );
        
        // Categorize by role
        const roleCount = {};
        owners.forEach(o => {
            roleCount[o.role] = (roleCount[o.role] || 0) + 1;
        });
        
        dataSourcesData.push({ name: sourceName, owners, roleCount });
    });
    
    // Sort by owner count
    dataSourcesData.sort((a, b) => b.owners.length - a.owners.length);
    
    container.innerHTML = dataSourcesData.map(dataSource => `
        <div class="data-source-card" onclick='showDataSourceDetails(${JSON.stringify(dataSource).replace(/'/g, "&apos;")})'>
            <div class="card-header">
                <div class="card-icon" style="background: linear-gradient(135deg, #667eea, #764ba2);">
                    <i class="fas fa-database"></i>
                </div>
                <div>
                    <div class="card-title">${dataSource.name}</div>
                    <div class="card-subtitle">${dataSource.owners.length} owners</div>
                </div>
            </div>
            <div class="role-distribution">
                ${Object.entries(dataSource.roleCount).slice(0, 3).map(([role, count]) => 
                    `<span class="role-badge">${role} (${count})</span>`
                ).join('')}
            </div>
            <div class="member-list">
                ${dataSource.owners.slice(0, 3).map(o => 
                    `<span class="member-chip">${o.name}</span>`
                ).join('')}
                ${dataSource.owners.length > 3 ? `<span class="member-chip">+${dataSource.owners.length - 3} more</span>` : ''}
            </div>
        </div>
    `).join('');
}

// ================================
// SERVICES VIEW
// ================================
function renderServicesView() {
    const container = document.getElementById('servicesGrid');
    
    if (AppState.services.size === 0) {
        container.innerHTML = '<p style="grid-column: 1/-1; text-align: center; color: var(--gray-500);">No services found</p>';
        return;
    }
    
    const servicesData = [];
    AppState.services.forEach(serviceName => {
        const admins = AppState.filteredData.filter(person => 
            person.services && person.services.split(';').map(s => s.trim()).includes(serviceName)
        );
        servicesData.push({ name: serviceName, admins });
    });
    
    container.innerHTML = servicesData.map(service => `
        <div class="service-card">
            <div class="card-header">
                <div class="card-icon">
                    <i class="fas fa-server"></i>
                </div>
                <div>
                    <div class="card-title">${service.name}</div>
                    <div class="card-subtitle">${service.admins.length} administrators</div>
                </div>
            </div>
            <div class="member-list">
                ${service.admins.length > 0 ? 
                    service.admins.map(a => 
                        `<span class="member-chip" onclick='showPersonDetails(${JSON.stringify(a).replace(/'/g, "&apos;")})'>${a.name}</span>`
                    ).join('') : 
                    '<p style="color: var(--gray-500);">No administrators assigned</p>'
                }
            </div>
        </div>
    `).join('');
}

// ================================
// FILTERING & SEARCH
// ================================
function populateFilters() {
    // Team filter
    const teamFilter = document.getElementById('teamFilter');
    teamFilter.innerHTML = '<option value="">All Teams</option>';
    Array.from(AppState.teams).sort().forEach(team => {
        teamFilter.innerHTML += `<option value="${team}">${team}</option>`;
    });
    
    // Institution filter
    const instFilter = document.getElementById('institutionFilter');
    instFilter.innerHTML = '<option value="">All Institutions</option>';
    Array.from(AppState.institutions).sort().forEach(inst => {
        instFilter.innerHTML += `<option value="${inst}">${inst}</option>`;
    });
    
    // Project filter
    const projFilter = document.getElementById('projectFilter');
    projFilter.innerHTML = '<option value="">All Projects</option>';
    Array.from(AppState.projects).sort().forEach(proj => {
        projFilter.innerHTML += `<option value="${proj}">${proj}</option>`;
    });
    
    // Data Source filter
    const dataSourceFilter = document.getElementById('dataSourceFilter');
    dataSourceFilter.innerHTML = '<option value="">All Data Sources</option>';
    Array.from(AppState.dataSources).sort().forEach(source => {
        dataSourceFilter.innerHTML += `<option value="${source}">${source}</option>`;
    });
}

function applyFilters() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const teamFilter = document.getElementById('teamFilter').value;
    const instFilter = document.getElementById('institutionFilter').value;
    const projFilter = document.getElementById('projectFilter').value;
    const dataSourceFilter = document.getElementById('dataSourceFilter').value;
    const statusFilter = document.getElementById('statusFilter').value;
    
    AppState.filteredData = AppState.rawData.filter(person => {
        // Search filter
        if (searchTerm) {
            const searchableFields = [
                person.name,
                person.email,
                person.role,
                person.team,
                person.institution,
                person.projects,
                person.data_sources,
                person.services
            ].filter(Boolean).join(' ').toLowerCase();
            
            if (!searchableFields.includes(searchTerm)) {
                return false;
            }
        }
        
        // Team filter
        if (teamFilter && (!person.team || !person.team.split(';').map(t => t.trim()).includes(teamFilter))) {
            return false;
        }
        
        // Institution filter
        if (instFilter && person.institution !== instFilter) {
            return false;
        }
        
        // Project filter
        if (projFilter && (!person.projects || !person.projects.split(';').map(p => p.trim()).includes(projFilter))) {
            return false;
        }
        
        // Data Source filter
        if (dataSourceFilter && (!person.data_sources || !person.data_sources.split(';').map(d => d.trim()).includes(dataSourceFilter))) {
            return false;
        }
        
        // Status filter
        if (statusFilter && (person.status || 'active') !== statusFilter) {
            return false;
        }
        
        return true;
    });
    
    AppState.currentPage = 1;
    renderCurrentView();
}

function clearFilters() {
    document.getElementById('searchInput').value = '';
    document.getElementById('teamFilter').value = '';
    document.getElementById('institutionFilter').value = '';
    document.getElementById('projectFilter').value = '';
    document.getElementById('dataSourceFilter').value = '';
    document.getElementById('statusFilter').value = '';
    
    AppState.filteredData = AppState.rawData;
    AppState.currentPage = 1;
    renderCurrentView();
}

// ================================
// SORTING
// ================================
function sortTable(column) {
    if (AppState.sortColumn === column) {
        AppState.sortDirection = AppState.sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        AppState.sortColumn = column;
        AppState.sortDirection = 'asc';
    }
    
    AppState.filteredData.sort((a, b) => {
        const aValue = a[column] || '';
        const bValue = b[column] || '';
        
        if (AppState.sortDirection === 'asc') {
            return aValue.localeCompare(bValue);
        } else {
            return bValue.localeCompare(aValue);
        }
    });
    
    renderTableView();
}

// ================================
// MODALS
// ================================
function showPersonDetails(person) {
    const modal = document.getElementById('personModal');
    const modalTitle = document.getElementById('modalPersonName');
    const modalBody = document.getElementById('modalBody');
    
    modalTitle.textContent = person.name;
    
    const teams = person.team ? person.team.split(';').map(t => `<span class="tag tag-team">${t.trim()}</span>`).join('') : 'Not assigned';
    const projects = person.projects ? person.projects.split(';').map(p => `<span class="tag tag-project">${p.trim()}</span>`).join('') : 'None';
    const dataSources = person.data_sources ? person.data_sources.split(';').map(d => `<span class="tag tag-data">${d.trim()}</span>`).join('') : 'None';
    const services = person.services ? person.services.split(';').map(s => `<span class="tag tag-service">${s.trim()}</span>`).join('') : 'None';
    
    modalBody.innerHTML = `
        <div class="person-details-modal">
            <div class="detail-row">
                <span class="detail-label">Email:</span>
                <span class="detail-value"><a href="mailto:${person.email}">${person.email}</a></span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Role:</span>
                <span class="detail-value">${person.role}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Institution:</span>
                <span class="detail-value">${person.institution || 'Not specified'}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Teams:</span>
                <span class="detail-value">${teams}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Projects:</span>
                <span class="detail-value">${projects}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Data Sources:</span>
                <span class="detail-value">${dataSources}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Services:</span>
                <span class="detail-value">${services}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label">Status:</span>
                <span class="detail-value">
                    <span class="status-badge ${person.status || 'active'}">${person.status || 'active'}</span>
                </span>
            </div>
        </div>
    `;
    
    modal.classList.add('active');
}

function showProjectDetails(project) {
    const modal = document.getElementById('projectModal');
    const modalTitle = document.getElementById('projectModalTitle');
    const modalBody = document.getElementById('projectModalBody');
    
    modalTitle.textContent = `Project: ${project.name}`;
    
    modalBody.innerHTML = `
        <div class="project-details">
            <h3>Team Members (${project.members.length})</h3>
            <div class="institution-breakdown">
                ${Object.entries(project.institutionCount).map(([inst, count]) => 
                    `<span class="inst-badge">${inst}: ${count}</span>`
                ).join('')}
            </div>
            <div class="project-team-list">
                ${project.members.map(m => `
                    <div class="project-member-card">
                        <div class="member-info">
                            <strong>${m.name}</strong> - ${m.role}
                            <br><small>${m.institution || 'No institution'}</small>
                            <br><small>${m.email}</small>
                        </div>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
    
    modal.classList.add('active');
}

function showDataSourceDetails(dataSource) {
    const modal = document.getElementById('dataSourceModal');
    const modalTitle = document.getElementById('dataSourceModalTitle');
    const modalBody = document.getElementById('dataSourceModalBody');
    
    modalTitle.textContent = `Data Source: ${dataSource.name}`;
    
    modalBody.innerHTML = `
        <div class="data-source-details">
            <h3>Data Owners (${dataSource.owners.length})</h3>
            <div class="role-breakdown">
                ${Object.entries(dataSource.roleCount).map(([role, count]) => 
                    `<span class="role-badge">${role}: ${count}</span>`
                ).join('')}
            </div>
            <div class="data-source-owners-list">
                ${dataSource.owners.map(o => `
                    <div class="owner-card">
                        <div class="owner-info">
                            <strong>${o.name}</strong> - ${o.role}
                            <br><small>${o.institution || 'No institution'}</small>
                            <br><small>${o.email}</small>
                        </div>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
    
    modal.classList.add('active');
}

function closePersonModal() {
    document.getElementById('personModal').classList.remove('active');
}

function closeProjectModal() {
    document.getElementById('projectModal').classList.remove('active');
}

function closeDataSourceModal() {
    document.getElementById('dataSourceModal').classList.remove('active');
}

// ================================
// EXPORT FUNCTIONALITY
// ================================
function exportAsCSV() {
    const headers = AppState.csvHeaders.length > 0 ? AppState.csvHeaders : 
        ['name', 'email', 'role', 'team', 'institution', 'projects', 'data_sources', 'services', 'status'];
    
    const rows = [headers.join(',')];
    
    AppState.filteredData.forEach(person => {
        const values = headers.map(header => {
            const fieldName = header.toLowerCase().replace(/\s+/g, '_');
            const value = person[fieldName] || '';
            return value.includes(',') ? `"${value}"` : value;
        });
        rows.push(values.join(','));
    });
    
    const csvContent = rows.join('\n');
    downloadFile(csvContent, 'personnel_export.csv', 'text/csv');
    showToast('Data exported as CSV', 'success');
}

function exportAsJSON() {
    const jsonContent = JSON.stringify(AppState.filteredData, null, 2);
    downloadFile(jsonContent, 'personnel_export.json', 'application/json');
    showToast('Data exported as JSON', 'success');
}

function exportAsPDF() {
    window.print();
    showToast('Use browser print dialog to save as PDF', 'info');
}

function printReport() {
    window.print();
}

function downloadFile(content, filename, type) {
    const blob = new Blob([content], { type: type });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    
    URL.revokeObjectURL(url);
}

// ================================
// LOCAL STORAGE
// ================================
function saveToLocalStorage() {
    const dataToSave = {
        rawData: AppState.rawData,
        csvHeaders: AppState.csvHeaders,
        timestamp: new Date().toISOString()
    };
    localStorage.setItem('orgChartData', JSON.stringify(dataToSave));
    localStorage.setItem('orgChartLastUpdate', new Date().toISOString());
}

// ================================
// UTILITY FUNCTIONS
// ================================
function getInitials(name) {
    if (!name) return '??';
    const parts = name.split(' ');
    if (parts.length >= 2) {
        return parts[0][0].toUpperCase() + parts[parts.length - 1][0].toUpperCase();
    }
    return name.substring(0, 2).toUpperCase();
}

function getTimeAgo(date) {
    const now = new Date();
    const diff = Math.floor((now - date) / 1000); // difference in seconds
    
    if (diff < 60) return 'just now';
    if (diff < 3600) return `${Math.floor(diff / 60)} minutes ago`;
    if (diff < 86400) return `${Math.floor(diff / 3600)} hours ago`;
    if (diff < 2592000) return `${Math.floor(diff / 86400)} days ago`;
    if (diff < 31536000) return `${Math.floor(diff / 2592000)} months ago`;
    return `${Math.floor(diff / 31536000)} years ago`;
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    const icons = {
        success: 'check-circle',
        error: 'exclamation-circle',
        warning: 'exclamation-triangle',
        info: 'info-circle'
    };
    
    toast.innerHTML = `
        <i class="fas fa-${icons[type]}"></i>
        <span>${message}</span>
    `;
    
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

function showLoadingSpinner(show) {
    const spinner = document.getElementById('loadingSpinner');
    if (spinner) {
        if (show) {
            spinner.classList.add('active');
        } else {
            spinner.classList.remove('active');
        }
    }
}

function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// ================================
// GLOBAL FUNCTION EXPORTS
// ================================
window.showPersonDetails = showPersonDetails;
window.showProjectDetails = showProjectDetails;
window.showDataSourceDetails = showDataSourceDetails;
window.changePage = changePage;
window.sortTable = sortTable;
