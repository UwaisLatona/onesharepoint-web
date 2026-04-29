// =============================================================================
// app.js — OneSharePoint Hub Provisioning Tool
// =============================================================================

// ===== STATE =====
let currentStep = 1;
const totalSteps = 4;
let selectedPages = [];
let selectedLibraries = [];
let userEmail = '';
let subPageCounters = { 1: 1, 2: 1, 3: 1 };


// =============================================================================
// MSAL AUTHENTICATION
// =============================================================================

const msalConfig = {
    auth: {
        clientId: '474b0cf9-1a01-4e65-907a-29c80c909c82',
        authority: 'https://login.microsoftonline.com/531ad315-a282-482e-9136-c89692ae41e8',
        redirectUri: window.location.origin
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginScopes = {
    scopes: ['https://onesharepoint.sharepoint.com/.default']
};

async function login() {
    try {
        const response = await msalInstance.loginPopup(loginScopes);

        // Extract user info from token
        userEmail = response.account.username;
        const displayName = response.account.name || userEmail;
        const initials = displayName
            .split(' ')
            .map(function (n) { return n[0]; })
            .join('')
            .toUpperCase();

        document.getElementById('userName').textContent = displayName;
        document.getElementById('userInitials').textContent = initials;
        document.getElementById('loginScreen').classList.add('hidden');
        document.getElementById('appShell').classList.remove('hidden');
    } catch (error) {
        console.error('Login failed:', error);
        alert('Sign-in failed. Please try again.');
    }
}

async function getAccessToken() {
    var accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
        throw new Error('No signed-in user');
    }

    try {
        // Try silent token acquisition first
        var response = await msalInstance.acquireTokenSilent({
            scopes: ['https://onesharepoint.sharepoint.com/.default'],
            account: accounts[0]
        });
        return response.accessToken;
    } catch (error) {
        // Fall back to popup if silent fails
        var popupResponse = await msalInstance.acquireTokenPopup({
            scopes: ['https://onesharepoint.sharepoint.com/.default']
        });
        return popupResponse.accessToken;
    }
}


// =============================================================================
// STEPPER NAVIGATION
// =============================================================================

function goToStep(step) {
    if (step > currentStep + 1) return; // prevent skipping ahead
    currentStep = step;
    renderStep();
}

function nextStep() {
    if (!validateStep(currentStep)) return;
    if (currentStep < totalSteps) {
        currentStep++;
        renderStep();
    }
}

function prevStep() {
    if (currentStep > 1) {
        currentStep--;
        renderStep();
    }
}

function renderStep() {
    // Show active section
    var sections = document.querySelectorAll('.form-section');
    for (var i = 0; i < sections.length; i++) {
        sections[i].classList.remove('active');
    }
    document.getElementById('step' + currentStep).classList.add('active');

    // Update stepper indicators
    var steps = document.querySelectorAll('.step');
    for (var j = 0; j < steps.length; j++) {
        var n = parseInt(steps[j].dataset.step);
        steps[j].classList.remove('active', 'completed');
        if (n === currentStep) steps[j].classList.add('active');
        else if (n < currentStep) steps[j].classList.add('completed');
    }

    // Update connectors
    for (var k = 1; k < totalSteps; k++) {
        var conn = document.getElementById('conn' + k);
        if (conn) {
            if (k < currentStep) {
                conn.classList.add('completed');
            } else {
                conn.classList.remove('completed');
            }
        }
    }

    // Populate review on step 4
    if (currentStep === 4) populateReview();

    updateSummary();
    window.scrollTo({ top: 0, behavior: 'smooth' });
}


// =============================================================================
// VALIDATION
// =============================================================================

function validateStep(step) {
    clearErrors();
    var valid = true;

    if (step === 1) {
        var area = document.getElementById('businessArea').value;
        var name = document.getElementById('hubName').value.trim();
        var sensitivity = document.getElementById('sensitivityLabel').value;
        var nameRegex = /^[a-zA-Z0-9\s\-]+$/;

        if (!area) {
            showError('businessArea', 'businessAreaError');
            valid = false;
        }
        if (!name || !nameRegex.test(name)) {
            showError('hubName', 'hubNameError');
            valid = false;
        }
        if (!sensitivity) {
            showError('sensitivityLabel', 'sensitivityError');
            valid = false;
        }
    }

    if (step === 3) {
        var theme = document.getElementById('hubTheme').value;
        if (!theme) {
            showError('hubTheme', 'themeError');
            valid = false;
        }
    }

    return valid;
}

function showError(fieldId, errorId) {
    document.getElementById(fieldId).classList.add('field-error');
    document.getElementById(errorId).classList.remove('hidden');
}

function clearErrors() {
    var errorFields = document.querySelectorAll('.field-error');
    for (var i = 0; i < errorFields.length; i++) {
        errorFields[i].classList.remove('field-error');
    }
    var errorTexts = document.querySelectorAll('.error-text');
    for (var j = 0; j < errorTexts.length; j++) {
        errorTexts[j].classList.add('hidden');
    }
}


// =============================================================================
// CARD TOGGLES (pages & libraries)
// =============================================================================

function toggleCard(el, type) {
    el.classList.toggle('selected');
    var value = el.dataset.value;

    if (type === 'pages') {
        if (el.classList.contains('selected')) {
            selectedPages.push(value);
        } else {
            selectedPages = selectedPages.filter(function (p) { return p !== value; });
        }
    } else {
        if (el.classList.contains('selected')) {
            selectedLibraries.push(value);
        } else {
            selectedLibraries = selectedLibraries.filter(function (l) { return l !== value; });
        }
    }

    updateSummary();
}


// =============================================================================
// CUSTOM PAGES — SUB-PAGE LOGIC
// =============================================================================

function addSubPage(containerId, pageNum) {
    if (subPageCounters[pageNum] >= 3) return; // max 3 sub-pages per custom page
    subPageCounters[pageNum]++;

    var container = document.getElementById(containerId);
    var input = document.createElement('input');
    input.type = 'text';
    input.className = 'sub-page-input';
    input.placeholder = 'Sub page (optional)';
    input.id = 'customPage' + pageNum + 'Sub' + subPageCounters[pageNum];
    input.style.marginTop = '6px';
    container.appendChild(input);
}

function buildCustomPagesJson() {
    var pages = [];
    for (var i = 1; i <= 3; i++) {
        var name = document.getElementById('customPage' + i).value.trim();
        if (!name) continue;

        var subPages = [];
        for (var j = 1; j <= 3; j++) {
            var subEl = document.getElementById('customPage' + i + 'Sub' + j);
            if (subEl && subEl.value.trim()) {
                subPages.push(subEl.value.trim());
            }
        }
        pages.push({ name: name, subPages: subPages });
    }
    return pages;
}


// =============================================================================
// SUMMARY PANEL
// =============================================================================

function updateSummary() {
    var area = document.getElementById('businessArea').value;
    var name = document.getElementById('hubName').value;
    var displayName = (area && name) ? area + ' - ' + name : (name || '\u2014');
    document.getElementById('sumName').textContent = displayName;
    document.getElementById('sumSensitivity').textContent =
        document.getElementById('sensitivityLabel').value || '\u2014';
    document.getElementById('sumTheme').textContent =
        document.getElementById('hubTheme').value || '\u2014';
    document.getElementById('sumPages').textContent =
        selectedPages.length ? selectedPages.length + ' selected' : '0 selected';
    document.getElementById('sumLibraries').textContent =
        selectedLibraries.length ? selectedLibraries.length + ' selected' : '0 selected';

    // Count custom pages
    var customCount = 0;
    for (var i = 1; i <= 3; i++) {
        if (document.getElementById('customPage' + i).value.trim()) customCount++;
    }
    document.getElementById('sumCustom').textContent =
        customCount ? customCount + ' added' : '0 added';
}

// Live-update summary as the user types
document.addEventListener('input', updateSummary);


// =============================================================================
// REVIEW STEP
// =============================================================================

function populateReview() {
    var customPagesJson = buildCustomPagesJson();
    var customNames = customPagesJson.map(function (p) { return p.name; }).join(', ');

    var html =
        '<div style="display: grid; grid-template-columns: 160px 1fr; gap: 8px 16px;">' +
            '<strong style="color: var(--text-muted);">Site name</strong>' +
            '<span>' + document.getElementById('hubName').value + '</span>' +
            '<strong style="color: var(--text-muted);">Sensitivity</strong>' +
            '<span>' + document.getElementById('sensitivityLabel').value + '</span>' +
            '<strong style="color: var(--text-muted);">Theme</strong>' +
            '<span>' + document.getElementById('hubTheme').value + '</span>' +
            '<strong style="color: var(--text-muted);">Pages</strong>' +
            '<span>' + (selectedPages.length ? selectedPages.join(', ') : 'None') + '</span>' +
            '<strong style="color: var(--text-muted);">Libraries</strong>' +
            '<span>' + (selectedLibraries.length ? selectedLibraries.join(', ') : 'None') + '</span>' +
            '<strong style="color: var(--text-muted);">Custom pages</strong>' +
            '<span>' + (customNames || 'None') + '</span>' +
            '<strong style="color: var(--text-muted);">Requested by</strong>' +
            '<span>' + userEmail + '</span>' +
        '</div>';

    document.getElementById('reviewContent').innerHTML = html;
}


// =============================================================================
// SUBMIT TO SHAREPOINT
// =============================================================================

async function submitRequest() {
    var submitBtn = document.getElementById('submitBtn');
    submitBtn.disabled = true;
    submitBtn.textContent = 'Submitting...';

    try {
        var token = await getAccessToken();

        var payload = {
            '__metadata': { 'type': 'SP.Data.Hub_x0020_RequestsListItem' },
            'Title': document.getElementById('businessArea').value + ' - ' + document.getElementById('hubName').value.trim(),
            'HubName': document.getElementById('businessArea').value + ' - ' + document.getElementById('hubName').value.trim(),
            'BusinessArea': document.getElementById('businessArea').value,
            'HubOwnerEmail': userEmail,
            'SensitivityLabel': document.getElementById('sensitivityLabel').value,
            'HubTheme': document.getElementById('hubTheme').value,
            'SelectedPages': selectedPages.join(', '),
            'SelectedLibraries': selectedLibraries.join(', '),
            'CustomPagesJson': JSON.stringify(buildCustomPagesJson()),
            'RequestStatus': 'Submitted'
        };

        // Check for duplicate in the list
        var checkUrl = "https://onesharepoint.sharepoint.com/sites/onesharepoint-admin/_api/web/lists/getbytitle('Hub Requests')/items" +
            "?$filter=HubName eq '" + payload.HubName + "'" +
            "&$select=Id&$top=1";

        var checkResponse = await fetch(checkUrl, {
            method: 'GET',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Accept': 'application/json;odata=verbose'
            }
        });

        var checkData = await checkResponse.json();
        if (checkData.d.results.length > 0) {
            alert('A hub with this name already exists. Please choose a different name.');
            return;
        }

        var response = await fetch(
            "https://onesharepoint.sharepoint.com/sites/onesharepoint-admin/_api/web/lists/getbytitle('Hub Requests')/items",
            {
                method: 'POST',
                headers: {
                    'Authorization': 'Bearer ' + token,
                    'Content-Type': 'application/json;odata=verbose',
                    'Accept': 'application/json;odata=verbose'
                },
                body: JSON.stringify(payload)
            }
        );

        if (!response.ok) {
            var errorData = await response.json();
            throw new Error(errorData.error?.message?.value || 'Submission failed');
        }

        // Success — show toast
        var toast = document.getElementById('toast');
        toast.classList.add('show');
        setTimeout(function () { toast.classList.remove('show'); }, 4000);

    } catch (error) {
        console.error('Submit failed:', error);
        alert('Submission failed: ' + error.message);
    } finally {
        submitBtn.disabled = false;
        submitBtn.innerHTML =
            '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
            '<path d="M22 2L11 13M22 2l-7 20-4-9-9-4 20-7z"/></svg> Send for approval';
    }
}


// =============================================================================
// PAGE NAVIGATION (New Request / My Requests)
// =============================================================================

function showPage(page) {
    var newRequestPage = document.getElementById('newRequestPage');
    var myRequestsPage = document.getElementById('myRequestsPage');
    var navNewRequest = document.getElementById('navNewRequest');
    var navMyRequests = document.getElementById('navMyRequests');
    var breadcrumb = document.querySelector('.topnav-breadcrumb');

    if (page === 'newRequest') {
        newRequestPage.classList.remove('hidden');
        myRequestsPage.classList.add('hidden');
        navNewRequest.classList.add('active');
        navMyRequests.classList.remove('active');
        breadcrumb.textContent = 'Hub Request';
    } else {
        newRequestPage.classList.add('hidden');
        myRequestsPage.classList.remove('hidden');
        navNewRequest.classList.remove('active');
        navMyRequests.classList.add('active');
        breadcrumb.textContent = 'My Requests';
        loadMyRequests();
    }
}


// =============================================================================
// MY REQUESTS — FETCH & DISPLAY
// =============================================================================

async function loadMyRequests() {
    var loadingEl = document.getElementById('requestsLoading');
    var emptyEl = document.getElementById('requestsEmpty');
    var tableEl = document.getElementById('requestsTable');
    var bodyEl = document.getElementById('requestsBody');

    // Reset states
    loadingEl.classList.remove('hidden');
    emptyEl.classList.add('hidden');
    tableEl.classList.add('hidden');
    bodyEl.innerHTML = '';

    try {
        var token = await getAccessToken();

        // Filter by current user's email
        var filterUrl = "https://onesharepoint.sharepoint.com/sites/onesharepoint-admin/_api/web/lists/getbytitle('Hub Requests')/items" +
            "?$filter=HubOwnerEmail eq '" + userEmail + "'" +
            "&$orderby=Created desc" +
            "&$select=BusinessArea,HubName,SensitivityLabel,HubTheme,RequestStatus,Created,ProvisionedUrl";

        var response = await fetch(filterUrl, {
            method: 'GET',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Accept': 'application/json;odata=verbose'
            }
        });

        if (!response.ok) {
            throw new Error('Failed to load requests');
        }

        var data = await response.json();
        var items = data.d.results;

        loadingEl.classList.add('hidden');

        if (items.length === 0) {
            emptyEl.classList.remove('hidden');
            return;
        }

        // Build table rows
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var row = document.createElement('tr');

            var statusClass = getStatusClass(item.RequestStatus);
            var dateStr = formatDate(item.Created);

            var hubNameCell = item.ProvisionedUrl
                ? '<a href="' + item.ProvisionedUrl + '" target="_blank" style="color:var(--accent);text-decoration:none;">' + item.HubName + '</a>'
                : item.HubName;

            row.innerHTML =
                '<td>' + (item.HubName || '\u2014') + '</td>' +
                '<td>' + (item.SensitivityLabel || '\u2014') + '</td>' +
                '<td>' + (item.HubTheme || '\u2014') + '</td>' +
                '<td><span class="status-badge ' + statusClass + '">' + (item.RequestStatus || 'Unknown') + '</span></td>' +
                '<td>' + dateStr + '</td>' +
                '<td>' + (item.ProvisionedUrl ? '<a href="' + item.ProvisionedUrl + '" target="_blank" style="color:var(--accent);text-decoration:none;">' + item.ProvisionedUrl + '</a>' : '\u2014') + '</td>';                
            bodyEl.appendChild(row);
        }

        tableEl.classList.remove('hidden');

    } catch (error) {
        console.error('Failed to load requests:', error);
        loadingEl.classList.add('hidden');
        emptyEl.classList.remove('hidden');
        emptyEl.querySelector('p').textContent = 'Failed to load requests. Please try again.';
    }
}

function getStatusClass(status) {
    if (!status) return 'status-pending';
    var s = status.toLowerCase();
    if (s === 'submitted') return 'status-submitted';
    if (s === 'approved') return 'status-approved';
    if (s === 'rejected') return 'status-rejected';
    if (s === 'provisioned' || s === 'completed') return 'status-provisioned';
    return 'status-pending';
}

function formatDate(dateStr) {
    if (!dateStr) return '\u2014';
    var date = new Date(dateStr);
    var day = date.getDate();
    var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var month = months[date.getMonth()];
    var year = date.getFullYear();
    var hours = date.getHours().toString().padStart(2, '0');
    var mins = date.getMinutes().toString().padStart(2, '0');
    return day + ' ' + month + ' ' + year + ' at ' + hours + ':' + mins;
}