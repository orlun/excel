import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, GoogleAuthProvider, signInWithPopup, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, collection, addDoc, updateDoc, deleteDoc, doc, onSnapshot, query, orderBy } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
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
const auth = getAuth(app);
const provider = new GoogleAuthProvider();

// Pon aquí tu correo para que SOLO TÚ puedas entrar
const ALLOWED_EMAIL = "ivan.sanchez.roman09@gmail.com"; 

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
const elAvailableBalance = document.getElementById('availableBalance');
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

const RENT_MONTHS = [
    { label: 'Jul 26', y: 2026, m: 6 },
    { label: 'Ago 26', y: 2026, m: 7 },
    { label: 'Sep 26', y: 2026, m: 8 },
    { label: 'Oct 26', y: 2026, m: 9 },
    { label: 'Nov 26', y: 2026, m: 10 },
    { label: 'Dic 26', y: 2026, m: 11 },
    { label: 'Ene 27', y: 2027, m: 0 },
    { label: 'Feb 27', y: 2027, m: 1 },
    { label: 'Mar 27', y: 2027, m: 2 },
    { label: 'Abr 27', y: 2027, m: 3 },
    { label: 'May 27', y: 2027, m: 4 },
    { label: 'Jun 27', y: 2027, m: 5 }
];

// ===== Auth & Initialization =====
onAuthStateChanged(auth, (user) => {
    const loginScreen = document.getElementById('loginScreen');
    const mainApp = document.getElementById('mainApp');
    const loginError = document.getElementById('loginError');

    if (user) {
        if (ALLOWED_EMAIL !== "tu_email@gmail.com" && user.email !== ALLOWED_EMAIL) {
            signOut(auth);
            loginError.textContent = `El correo ${user.email} no está autorizado.`;
            loginError.classList.remove('hidden');
            loginScreen.classList.remove('hidden');
            mainApp.classList.add('hidden');
        } else if (ALLOWED_EMAIL === "tu_email@gmail.com") {
            loginError.innerHTML = `⚠️ Estás dentro con ${user.email}, pero aún no has configurado tu correo en el código. Edita app.js y cambia ALLOWED_EMAIL por tu correo real.`;
            loginError.classList.remove('hidden');
            loginScreen.classList.add('hidden');
            mainApp.classList.remove('hidden');
            initAppOnce();
        } else {
            loginScreen.classList.add('hidden');
            mainApp.classList.remove('hidden');
            initAppOnce();
        }
    } else {
        loginScreen.classList.remove('hidden');
        mainApp.classList.add('hidden');
    }
});

let appInitialized = false;
function initAppOnce() {
    if (appInitialized) return;
    appInitialized = true;
    
    elDate.textContent = formatDate(new Date().toISOString());
    document.getElementById('date').valueAsDate = new Date();
    
    setupEventListeners();
    listenToFirebase();
}

document.getElementById('btnLogin').addEventListener('click', async () => {
    const loginError = document.getElementById('loginError');
    loginError.classList.add('hidden');
    try {
        await signInWithPopup(auth, provider);
    } catch (error) {
        console.error("Error al iniciar sesión:", error);
        loginError.textContent = "Error al iniciar sesión: " + error.message + " (Recuerda habilitar Google Sign-In en Firebase Console)";
        loginError.classList.remove('hidden');
    }
});

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
        if (!settings.geminiKey) {
            alert("⚠️ No has configurado la API Key. Ve a los ajustes (el icono del engranaje) y pon tu clave de Gemini.");
            return;
        }
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
        Elige category estrictamente entre: Comida supermercado, Comidas fuera, Ocio, Transporte, Facturas, Ropa, Piso, Otros.`;
        
        const image = {
            inlineData: { data: base64Data, mimeType: mimeType }
        };
        
        const result = await model.generateContent([prompt, image]);
        const responseText = result.response.text();
        
        let jsonStr = responseText;
        const match = responseText.match(/\{[\s\S]*\}/);
        if (match) {
            jsonStr = match[0];
        } else {
            jsonStr = responseText.replace(/```json/g, '').replace(/```/g, '').trim();
        }
        const data = JSON.parse(jsonStr);
        
        if (data.concept) document.getElementById('concept').value = data.concept;
        if (data.amount) document.getElementById('amount').value = data.amount;
        if (data.category) document.getElementById('category').value = data.category;
        if (data.type) {
            document.querySelector(`input[name="type"][value="${data.type}"]`).checked = true;
        }
    } catch(err) {
        console.error("AI Error:", err);
        alert("No se pudo analizar el ticket con la IA. Error: " + err.message + "\n\nComprueba que tu API Key está bien puesta en la ruedecilla de ajustes.");
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
        document.getElementById('editTransactionId').value = '';
        document.querySelector('#addModal h2').innerHTML = 'Nuevo Movimiento';
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

        const editId = document.getElementById('editTransactionId').value;
        
        // Si no subió foto nueva, mantenemos la anterior si está editando
        if (!photoUrl && editId) {
            const existingT = transactions.find(x => x.id === editId);
            if (existingT && ticketPreview.src === existingT.photo) {
                photoUrl = existingT.photo;
            }
        }

        const transactionData = {
            type,
            amount,
            concept,
            category,
            date,
            photo: photoUrl
        };

        if (editId) {
            await updateDoc(doc(db, "transactions", editId), transactionData);
        } else {
            transactionData.timestamp = new Date().getTime();
            await addDoc(collection(db, "transactions"), transactionData);
        }
        
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
            if (t.amount <= 6 && (t.category === 'Ocio' || t.category === 'Comida' || t.category === 'Comidas fuera' || t.category === 'Comida supermercado' || t.category === 'Otros')) {
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
    
    let categoryTotals = { 'Comida supermercado': 0, 'Comidas fuera': 0, 'Ocio': 0, 'Transporte': 0, 'Facturas': 0, 'Ropa': 0, 'Piso': 0, 'Otros': 0 };

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
    
    const endOfCourse = new Date('2027-07-31');
    let remainingDays = Math.max(1, Math.floor((endOfCourse - now) / (1000 * 60 * 60 * 24)));
    
    // Contar cuántos meses de alquiler quedan realmente pendientes verificando los pagos
    let rentMonthsPending = 0;
    RENT_MONTHS.forEach(rm => {
        // Julio 2026 se marca como pagado por defecto
        const isPaid = (rm.y === 2026 && rm.m === 6) || transactions.some(t => {
            const d = new Date(t.date);
            return t.type === 'gasto' && 
                   d.getFullYear() === rm.y && 
                   d.getMonth() === rm.m && 
                   (t.category === 'Piso' || t.concept.toLowerCase().includes('alquiler'));
        });
        if (!isPaid) rentMonthsPending++;
    });
    
    const availableForDaily = balance - (rentMonthsPending * settings.rent);
    let dailyBudget = availableForDaily > 0 ? (availableForDaily / remainingDays) : 0;

    const daysSinceStart = Math.max(1, Math.floor((now - start) / (1000 * 60 * 60 * 24)));
    const avgDailyExpense = totalExpense / daysSinceStart;
    
    let viabilityMonths = 0;
    if (avgDailyExpense > 0) {
        viabilityMonths = balance / (avgDailyExpense * 30.44);
    }

    return {
        balance,
        availableForDaily,
        monthExpense: monthExpenseStr,
        monthIncome: monthIncomeStr,
        dailyBudget,
        viabilityMonths,
        categoryTotals,
        rentMonthsPending
    };
}

// ===== Rendering =====
function renderDashboard() {
    const metrics = calculateMetrics();
    const engine = calculateAntigravityEngine();
    
    if (elAvailableBalance) elAvailableBalance.textContent = formatCurrency(metrics.availableForDaily);
    if (elTotalBalance) elTotalBalance.textContent = formatCurrency(metrics.balance);
    
    elMonthExpense.textContent = formatCurrency(metrics.monthExpense);
    elMonthIncome.textContent = formatCurrency(metrics.monthIncome);
    elDailyBudget.textContent = formatCurrency(metrics.dailyBudget);
    
    // Render Rent Tracking
    const rentList = document.getElementById('rentList');
    if (rentList) {
        rentList.innerHTML = '';
        RENT_MONTHS.forEach(rm => {
            // Julio 2026 se marca como pagado por defecto
            const isPaid = (rm.y === 2026 && rm.m === 6) || transactions.some(t => {
                const d = new Date(t.date);
                return t.type === 'gasto' && 
                   d.getFullYear() === rm.y && 
                   d.getMonth() === rm.m && 
                   (t.category === 'Piso' || t.concept.toLowerCase().includes('alquiler'));
            });
            
            const div = document.createElement('div');
            div.className = `rent-month ${isPaid ? 'paid' : 'pending'}`;
            div.innerHTML = `<span>${rm.label}</span> <i class="fa-solid ${isPaid ? 'fa-check' : 'fa-clock'}"></i>`;
            rentList.appendChild(div);
        });
        
        const rentPendingCount = document.getElementById('rentPendingCount');
        if (rentPendingCount) {
            rentPendingCount.textContent = `${metrics.rentMonthsPending} pendientes`;
        }
    }
    
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
        const catColors = {
            'Comida supermercado': '#10b981',
            'Comidas fuera': '#3b82f6',
            'Ocio': '#8b5cf6',
            'Ropa': '#ec4899',
            'Transporte': '#f59e0b',
            'Facturas': '#06b6d4',
            'Piso': '#6366f1',
            'Otros': '#64748b'
        };
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
            <div class="transaction-actions" style="margin-top: 8px; display: flex; gap: 12px; justify-content: flex-end;">
                <button onclick="window.editTransaction('${t.id}')" style="background:none; border:none; color: var(--text-secondary); cursor:pointer; font-size: 0.9rem;" title="Editar"><i class="fa-solid fa-pen"></i></button>
                <button onclick="window.deleteTransaction('${t.id}')" style="background:none; border:none; color: var(--danger); cursor:pointer; font-size: 0.9rem;" title="Borrar"><i class="fa-solid fa-trash"></i></button>
            </div>
        </div>
    `;
    return div;
}

window.editTransaction = function(id) {
    const t = transactions.find(x => x.id === id);
    if (!t) return;
    
    document.getElementById('editTransactionId').value = t.id;
    document.querySelector(`input[name="type"][value="${t.type}"]`).checked = true;
    document.getElementById('amount').value = t.amount;
    document.getElementById('concept').value = t.concept;
    document.getElementById('category').value = t.category || 'Otros';
    document.getElementById('date').value = t.date;
    
    if (t.photo) {
        ticketPreview.src = t.photo;
        ticketPreview.classList.remove('hidden');
    } else {
        ticketPreview.src = '';
        ticketPreview.classList.add('hidden');
    }
    
    document.querySelector('#addModal h2').innerHTML = '<i class="fa-solid fa-pen"></i> Editar Movimiento';
    openModal(addModal);
};

window.deleteTransaction = async function(id) {
    if(confirm('¿Estás seguro de que quieres borrar este gasto?')) {
        try {
            await deleteDoc(doc(db, "transactions", id));
        } catch(e) {
            alert('Error al borrar: ' + e.message);
        }
    }
};

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

// La inicialización la hace ahora onAuthStateChanged
// document.addEventListener('DOMContentLoaded', init);
