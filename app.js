import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getFirestore, collection, addDoc, onSnapshot, query, orderBy } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import { getStorage, ref, uploadString, getDownloadURL } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-storage.js";

// Tu configuración de Firebase
const firebaseConfig = {
  apiKey: "AIzaSyAZuXPpmXhNSaV8UofSVMJA4GISIkaPDY4",
  authDomain: "excel-d4f12.firebaseapp.com",
  databaseURL: "https://excel-d4f12-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "excel-d4f12",
  storageBucket: "excel-d4f12.firebasestorage.app",
  messagingSenderId: "206082369803",
  appId: "1:206082369803:web:acfdaeee29ff7613b59a0e",
  measurementId: "G-YSBV27Y079"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const storage = getStorage(app);

// ===== Constants & State =====
let transactions = [];
let settings = JSON.parse(localStorage.getItem('fin_settings')) || {
    capital: 6820,
    rent: 375,
    startDate: '2026-07-01'
};

// Date formatter
const formatDate = (dateString) => {
    const options = { day: '2-digit', month: 'short', year: 'numeric' };
    return new Date(dateString).toLocaleDateString('es-ES', options);
};

const formatCurrency = (amount) => {
    return new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR' }).format(amount);
};

// ===== DOM Elements =====
const navItems = document.querySelectorAll('.bottom-nav .nav-item[data-tab]');
const tabContents = document.querySelectorAll('.tab-content');

const elDate = document.getElementById('currentDate');
const elTotalBalance = document.getElementById('totalBalance');
const elMonthExpense = document.getElementById('monthExpense');
const elMonthIncome = document.getElementById('monthIncome');
const elDailyBudget = document.getElementById('dailyBudget');
const elViability = document.getElementById('viability');
const elRecentTransactions = document.getElementById('recentTransactionsList');

const elAllTransactions = document.getElementById('allTransactionsList');
const elMonthFilter = document.getElementById('monthFilter');
const btnExport = document.getElementById('btnExport');

const addModal = document.getElementById('addModal');
const settingsModal = document.getElementById('settingsModal');
const viewTicketModal = document.getElementById('viewTicketModal');

const addForm = document.getElementById('addForm');
const settingsForm = document.getElementById('settingsForm');
const ticketPhoto = document.getElementById('ticketPhoto');
const ticketPreview = document.getElementById('ticketPreview');
const fullTicketImage = document.getElementById('fullTicketImage');

// Add a loading state for the submit button
const submitBtn = addForm.querySelector('button[type="submit"]');

// ===== Initialization =====
function init() {
    elDate.textContent = formatDate(new Date().toISOString());
    document.getElementById('date').valueAsDate = new Date();
    
    setupEventListeners();
    listenToFirebase();
}

function listenToFirebase() {
    const q = query(collection(db, "transactions"), orderBy("timestamp", "desc"));
    
    onSnapshot(q, (snapshot) => {
        transactions = [];
        snapshot.forEach((doc) => {
            transactions.push({ id: doc.id, ...doc.data() });
        });
        
        renderDashboard();
        populateMonthFilter();
        renderHistory();
    });
}

function setupEventListeners() {
    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            navItems.forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            
            tabContents.forEach(tc => tc.classList.remove('active'));
            document.getElementById(`tab-${item.dataset.tab}`).classList.add('active');
            
            if(item.dataset.tab === 'history') renderHistory();
        });
    });

    document.getElementById('btnAddExpense').addEventListener('click', (e) => {
        e.preventDefault();
        openModal(addModal);
    });
    
    document.getElementById('btnSettings').addEventListener('click', () => {
        document.getElementById('setCapital').value = settings.capital;
        document.getElementById('setRent').value = settings.rent;
        openModal(settingsModal);
    });

    document.querySelectorAll('.close-modal').forEach(btn => {
        btn.addEventListener('click', (e) => {
            closeModal(e.target.closest('.modal-overlay'));
        });
    });

    ticketPhoto.addEventListener('change', function() {
        const file = this.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                ticketPreview.src = e.target.result;
                ticketPreview.classList.remove('hidden');
            }
            reader.readAsDataURL(file);
        } else {
            ticketPreview.classList.add('hidden');
            ticketPreview.src = '';
        }
    });

    addForm.addEventListener('submit', handleAddTransaction);
    settingsForm.addEventListener('submit', handleSaveSettings);
    elMonthFilter.addEventListener('change', renderHistory);
    btnExport.addEventListener('click', exportCSV);
    
    document.getElementById('btnResetData').addEventListener('click', () => {
        alert('Para borrar datos ahora debes hacerlo desde Firebase Console (Firestore).');
    });
}

function openModal(modal) { modal.classList.add('active'); }
function closeModal(modal) {
    modal.classList.remove('active');
    if(modal === addModal) {
        addForm.reset();
        document.getElementById('date').valueAsDate = new Date();
        ticketPreview.classList.add('hidden');
        ticketPreview.src = '';
    }
}

function saveSettings() {
    localStorage.setItem('fin_settings', JSON.stringify(settings));
}

function handleSaveSettings(e) {
    e.preventDefault();
    settings.capital = parseFloat(document.getElementById('setCapital').value);
    settings.rent = parseFloat(document.getElementById('setRent').value);
    saveSettings();
    closeModal(settingsModal);
    renderDashboard();
}

async function handleAddTransaction(e) {
    e.preventDefault();
    submitBtn.disabled = true;
    submitBtn.textContent = 'Guardando en la nube...';
    
    try {
        const type = document.querySelector('input[name="type"]:checked').value;
        const amount = parseFloat(document.getElementById('amount').value);
        const concept = document.getElementById('concept').value;
        const date = document.getElementById('date').value;
        
        let photoUrl = null;
        
        // Si hay una foto cargada, subir a Storage
        if (ticketPreview.src && ticketPreview.src.startsWith('data:image')) {
            const storageRef = ref(storage, `tickets/${Date.now()}.jpg`);
            const snapshot = await uploadString(storageRef, ticketPreview.src, 'data_url');
            photoUrl = await getDownloadURL(snapshot.ref);
        }

        const transactionData = {
            type,
            amount,
            concept,
            date,
            photo: photoUrl,
            timestamp: new Date().getTime()
        };

        // Guardar en Firestore Database
        await addDoc(collection(db, "transactions"), transactionData);
        
        closeModal(addModal);
    } catch (error) {
        console.error("Error adding document: ", error);
        alert('Error al guardar: ' + error.message);
    } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = 'Guardar Movimiento';
    }
}

// ===== Calculations =====
function calculateMetrics() {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    let totalExpense = 0;
    let totalIncome = 0;
    let monthExpenseStr = 0;
    let monthIncomeStr = 0;

    transactions.forEach(t => {
        if (t.type === 'gasto') totalExpense += t.amount;
        if (t.type === 'ingreso') totalIncome += t.amount;
        
        const tDate = new Date(t.date);
        if (tDate.getMonth() === currentMonth && tDate.getFullYear() === currentYear) {
            if (t.type === 'gasto') monthExpenseStr += t.amount;
            if (t.type === 'ingreso') monthIncomeStr += t.amount;
        }
    });

    const start = new Date(settings.startDate);
    const balance = settings.capital + totalIncome - totalExpense;
    
    const endOfCourse = new Date('2027-06-30');
    let remainingDays = Math.max(1, Math.floor((endOfCourse - now) / (1000 * 60 * 60 * 24)));
    
    let endMonth = endOfCourse.getMonth();
    let endYear = endOfCourse.getFullYear();
    let remainingMonthsRent = (endYear - currentYear) * 12 + (endMonth - currentMonth);
    if (remainingMonthsRent < 0) remainingMonthsRent = 0;
    
    const availableForDaily = balance - (remainingMonthsRent * settings.rent);
    let dailyBudget = availableForDaily > 0 ? (availableForDaily / remainingDays) : 0;

    const daysSinceStart = Math.max(1, Math.floor((now - start) / (1000 * 60 * 60 * 24)));
    const avgDailyExpense = totalExpense / daysSinceStart;
    
    let viabilityMonths = 0;
    if (avgDailyExpense > 0) {
        viabilityMonths = balance / (avgDailyExpense * 30.44);
    }

    return {
        balance,
        monthExpense: monthExpenseStr,
        monthIncome: monthIncomeStr,
        dailyBudget,
        viabilityMonths
    };
}

// ===== Rendering =====
function renderDashboard() {
    const metrics = calculateMetrics();
    
    elTotalBalance.textContent = formatCurrency(metrics.balance);
    elMonthExpense.textContent = formatCurrency(metrics.monthExpense);
    elMonthIncome.textContent = formatCurrency(metrics.monthIncome);
    elDailyBudget.textContent = formatCurrency(metrics.dailyBudget);
    
    elViability.textContent = metrics.viabilityMonths > 99 ? '∞ meses' : `${metrics.viabilityMonths.toFixed(1)} meses`;
    if(metrics.viabilityMonths < 3 && metrics.balance > 0) {
        elViability.className = 'insight-value text-danger';
    } else {
        elViability.className = 'insight-value';
    }

    elRecentTransactions.innerHTML = '';
    const recent = transactions.slice(0, 5);
    
    if (recent.length === 0) {
        elRecentTransactions.innerHTML = '<p class="empty-state">No hay movimientos recientes.</p>';
        return;
    }

    recent.forEach(t => {
        elRecentTransactions.appendChild(createTransactionElement(t));
    });
}

function populateMonthFilter() {
    const months = new Set();
    transactions.forEach(t => {
        const d = new Date(t.date);
        const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        months.add(key);
    });
    
    elMonthFilter.innerHTML = '<option value="all">Todos los meses</option>';
    
    Array.from(months).sort().reverse().forEach(key => {
        const [year, month] = key.split('-');
        const date = new Date(year, parseInt(month) - 1, 1);
        const name = date.toLocaleDateString('es-ES', { month: 'long', year: 'numeric' });
        const capitalized = name.charAt(0).toUpperCase() + name.slice(1);
        
        elMonthFilter.innerHTML += `<option value="${key}">${capitalized}</option>`;
    });
}

function renderHistory() {
    const filter = elMonthFilter.value;
    elAllTransactions.innerHTML = '';
    
    let filtered = transactions;
    if (filter !== 'all') {
        filtered = transactions.filter(t => {
            const d = new Date(t.date);
            const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
            return key === filter;
        });
    }

    if (filtered.length === 0) {
        elAllTransactions.innerHTML = '<p class="empty-state">No hay movimientos en este periodo.</p>';
        return;
    }

    filtered.forEach(t => {
        elAllTransactions.appendChild(createTransactionElement(t));
    });
}

function createTransactionElement(t) {
    const div = document.createElement('div');
    div.className = 'transaction-item';
    
    const isExpense = t.type === 'gasto';
    const iconClass = isExpense ? 'fa-arrow-down' : 'fa-arrow-up';
    const amountClass = isExpense ? 'text-danger' : 'text-success';
    const sign = isExpense ? '-' : '+';
    
    let ticketHTML = '';
    if (t.photo) {
        ticketHTML = `<div class="has-ticket" onclick="viewTicket('${t.id}')">
            <i class="fa-solid fa-receipt"></i> Ver ticket
        </div>`;
    }

    div.innerHTML = `
        <div class="transaction-icon ${t.type}">
            <i class="fa-solid ${iconClass}"></i>
        </div>
        <div class="transaction-details">
            <div class="transaction-title">${t.concept}</div>
            <div class="transaction-date">${formatDate(t.date)}</div>
        </div>
        <div class="transaction-amount ${amountClass}">
            ${sign}${formatCurrency(t.amount)}
            ${ticketHTML}
        </div>
    `;
    return div;
}

window.viewTicket = function(id) {
    const t = transactions.find(x => x.id === id);
    if (t && t.photo) {
        fullTicketImage.src = t.photo;
        openModal(viewTicketModal);
    }
};

function exportCSV() {
    if(transactions.length === 0) return alert('No hay datos para exportar');
    
    let csv = 'Fecha,Concepto,Tipo,Importe,Ticket_URL\n';
    transactions.forEach(t => {
        csv += `${t.date},"${t.concept.replace(/"/g, '""')}",${t.type},${t.amount},"${t.photo || ''}"\n`;
    });
    
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', 'mis_cuentas.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

document.addEventListener('DOMContentLoaded', init);
