import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, collection, addDoc, onSnapshot, query, orderBy } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import { getStorage, ref, uploadString, getDownloadURL } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-storage.js";
import { GoogleGenerativeAI } from "https://esm.run/@google/generative-ai";

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

// Initialize Firebase con caché local para carga ultra rápida
const app = initializeApp(firebaseConfig);
const db = initializeFirestore(app, {
  localCache: persistentLocalCache({tabManager: persistentMultipleTabManager()})
});
const storage = getStorage(app);

// ===== Constants & State =====
let transactions = [];
let settings = JSON.parse(localStorage.getItem('fin_settings')) || {
    capital: 6820,
    rent: 375,
    startDate: '2026-07-01',
    geminiKey: ''
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
    try {
        const q = query(collection(db, "transactions"), orderBy("timestamp", "desc"));
        
        onSnapshot(q, { includeMetadataChanges: true }, (snapshot) => {
            try {
                transactions = [];
                snapshot.forEach((doc) => {
                    transactions.push({ id: doc.id, ...doc.data(), pending: doc.metadata.hasPendingWrites });
                });
                
                renderDashboard();
                populateMonthFilter();
                renderHistory();
            } catch (renderError) {
                console.error("Render Error:", renderError);
                alert("Error al dibujar la interfaz: " + renderError.message);
            }
        }, (error) => {
            console.error("Firebase Error:", error);
            alert("Error conectando a Firebase: " + error.code + " - " + error.message);
        });
    } catch (initError) {
        console.error("Init Error:", initError);
        alert("Error inicializando lectura: " + initError.message);
    }
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
        document.getElementById('setGeminiKey').value = settings.geminiKey || '';
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
            reader.onload = async function(e) {
                ticketPreview.src = e.target.result;
                ticketPreview.classList.remove('hidden');
                
                if (settings.geminiKey) {
                    await analyzeTicketWithAI(e.target.result);
                }
            }
            reader.readAsDataURL(file);
        } else {
            ticketPreview.classList.add('hidden');
            ticketPreview.src = '';
        }
    });

async function analyzeTicketWithAI(base64Str) {
    const aiStatus = document.getElementById('aiStatus');
    aiStatus.classList.remove('hidden');
    
    try {
        const genAI = new GoogleGenerativeAI(settings.geminiKey);
        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        
        const base64Data = base64Str.split(',')[1];
        const mimeType = base64Str.match(/data:(.*?);/)[1];
        
        const prompt = `Eres un asistente experto en finanzas que lee tickets. Extrae la información y devuelve SOLO un objeto JSON válido con este exacto formato, sin nada más:
        {
           "concept": "Resumen del lugar o compra (ej: Mercadona, Gasolinera Repsol)",
           "amount": 23.50,
           "type": "gasto",
           "category": "Comida"
        }
        Elige category estrictamente entre: Comida, Ocio, Transporte, Facturas, Otros.`;
        
        const image = {
            inlineData: { data: base64Data, mimeType: mimeType }
        };
        
        const result = await model.generateContent([prompt, image]);
        const responseText = result.response.text();
        
        const jsonStr = responseText.replace(/```json/g, '').replace(/```/g, '').trim();
        const data = JSON.parse(jsonStr);
        
        if (data.concept) document.getElementById('concept').value = data.concept;
        if (data.amount) document.getElementById('amount').value = data.amount;
        if (data.category) document.getElementById('category').value = data.category;
        if (data.type) {
            document.querySelector(`input[name="type"][value="${data.type}"]`).checked = true;
        }
    } catch(err) {
        console.error("AI Error:", err);
        alert("No se pudo analizar el ticket con la IA. Comprueba tu API Key.");
    } finally {
        aiStatus.classList.add('hidden');
    }
}

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
    settings.geminiKey = document.getElementById('setGeminiKey').value;
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
        const category = document.getElementById('category').value;
        
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
            category,
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
function calculateAntigravityEngine() {
    let streak = 0;
    const now = new Date();
    let checkDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    
    const expensesByDay = {};
    transactions.forEach(t => {
        if(t.type === 'gasto') {
            const d = new Date(t.date);
            const key = `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
            expensesByDay[key] = (expensesByDay[key] || 0) + t.amount;
        }
    });

    for(let i=0; i<30; i++) {
        const key = `${checkDate.getFullYear()}-${checkDate.getMonth()}-${checkDate.getDate()}`;
        if (!expensesByDay[key] || expensesByDay[key] === 0) {
            streak++;
        } else {
            break;
        }
        checkDate.setDate(checkDate.getDate() - 1);
    }

    let hormigaTotal = 0;
    let hormigaCount = 0;
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    transactions.forEach(t => {
        const d = new Date(t.date);
        if (t.type === 'gasto' && d.getMonth() === currentMonth && d.getFullYear() === currentYear) {
            if (t.amount <= 6 && (t.category === 'Ocio' || t.category === 'Comida' || t.category === 'Otros')) {
                hormigaTotal += t.amount;
                hormigaCount++;
            }
        }
    });

    return { streak, hormigaTotal, hormigaCount };
}

function calculateMetrics() {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    let totalExpense = 0;
    let totalIncome = 0;
    let monthExpenseStr = 0;
    let monthIncomeStr = 0;
    
    let categoryTotals = { 'Comida': 0, 'Ocio': 0, 'Transporte': 0, 'Facturas': 0, 'Otros': 0 };

    transactions.forEach(t => {
        if (t.type === 'gasto') totalExpense += t.amount;
        if (t.type === 'ingreso') totalIncome += t.amount;
        
        const tDate = new Date(t.date);
        if (tDate.getMonth() === currentMonth && tDate.getFullYear() === currentYear) {
            if (t.type === 'gasto') {
                monthExpenseStr += t.amount;
                let cat = t.category || 'Otros';
                if (!categoryTotals[cat]) categoryTotals[cat] = 0;
                categoryTotals[cat] += t.amount;
            }
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
        viabilityMonths,
        categoryTotals
    };
}

// ===== Rendering =====
function renderDashboard() {
    const metrics = calculateMetrics();
    const engine = calculateAntigravityEngine();
    
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

    const streakEl = document.getElementById('streakIndicator');
    if (streakEl) {
        if (engine.streak >= 1) {
            streakEl.classList.remove('hidden');
            document.getElementById('streakDays').textContent = engine.streak;
        } else {
            streakEl.classList.add('hidden');
        }
    }

    const alertsContainer = document.getElementById('antigravityAlerts');
    if (alertsContainer) {
        alertsContainer.innerHTML = '';
        if (engine.hormigaTotal > 20) {
            const diasSuper = (engine.hormigaTotal / (metrics.dailyBudget || 1)).toFixed(1);
            alertsContainer.innerHTML += `
                <div class="alert-card">
                    <i class="fa-solid fa-triangle-exclamation"></i>
                    <div class="alert-content">
                        <p><strong>Alerta Gasto Hormiga:</strong> ${formatCurrency(engine.hormigaTotal)} en ${engine.hormigaCount} micropagos.</p>
                        <p style="font-size: 0.75rem; opacity: 0.8; margin-top:4px;">Equivale a ${diasSuper} días completos de tu presupuesto. 🚨</p>
                    </div>
                </div>
            `;
        }
    }

    const catList = document.getElementById('categoryList');
    if (catList) {
        catList.innerHTML = '';
        const catColors = { 'Comida': '#10b981', 'Ocio': '#8b5cf6', 'Transporte': '#f59e0b', 'Facturas': '#ef4444', 'Otros': '#64748b' };
        const sortedCats = Object.entries(metrics.categoryTotals).sort((a,b) => b[1] - a[1]);
        
        sortedCats.forEach(([cat, amount]) => {
            if (amount > 0) {
                let percentage = metrics.monthExpense > 0 ? (amount / metrics.monthExpense) * 100 : 0;
                let color = catColors[cat] || '#64748b';
                catList.innerHTML += `
                    <div class="category-bar-container">
                        <div class="category-bar-header">
                            <span>${cat}</span>
                            <span>${formatCurrency(amount)}</span>
                        </div>
                        <div class="category-bar-bg">
                            <div class="category-bar-fill" style="width: ${percentage}%; background-color: ${color}"></div>
                        </div>
                    </div>
                `;
            }
        });
        if (metrics.monthExpense === 0) {
            catList.innerHTML = '<p class="empty-state">No hay gastos categorizados este mes.</p>';
        }
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
    
    let syncHTML = t.pending 
        ? `<i class="fa-solid fa-cloud-arrow-up sync-status" title="Sincronizando con la nube..."></i>` 
        : `<i class="fa-solid fa-cloud-check sync-status synced" title="Guardado en Firebase"></i>`;
    
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
            <div class="transaction-title">${t.concept} ${syncHTML}</div>
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
