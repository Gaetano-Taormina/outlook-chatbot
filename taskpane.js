Office.initialize = function () {};

let myMSALObj;
let accessToken = null;
let foundEmails = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        const msalConfig = {
            auth: { clientId: "65c3d535-927c-4930-b0e9-0cefddf82495", authority: "https://login.microsoftonline.com/common", redirectUri: "https://gaetano-taormina.github.io/outlook-chatbot/taskpane.html" },
            cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
        };
        myMSALObj = new msal.PublicClientApplication(msalConfig);

        document.getElementById("loginBtn").addEventListener("click", login);
        document.getElementById("actionSelector").addEventListener("change", aggiornaInterfaccia);
        document.getElementById("executeBtn").addEventListener("click", eseguiAzioneScelta);
        document.getElementById("selectAllBtn").addEventListener("click", toggleSelezionaTutte);
    }
});

function aggiornaStato(testo) {
    const statusEl = document.getElementById("status-text");
    if (statusEl) statusEl.innerText = testo;
}

function triggerAnimazione() {
    const wrapper = document.getElementById("inputWrapper");
    wrapper.classList.remove("animate-ui");
    void wrapper.offsetWidth; 
    wrapper.classList.add("animate-ui");
}

async function login() {
    try {
        const loginRequest = { scopes: ["user.read", "mail.readwrite"] };
        const loginResponse = await myMSALObj.loginPopup(loginRequest);
        
        const tokenRequest = { scopes: ["mail.readwrite"], account: loginResponse.account };
        const tokenResponse = await myMSALObj.acquireTokenSilent(tokenRequest);
        accessToken = tokenResponse.accessToken;
        
        document.getElementById("login-screen").style.display = "none";
        document.getElementById("app-screen").style.display = "flex";
        aggiornaStato("🟢 Connesso. Seleziona un'azione.");
    } catch (error) { alert("Errore di connessione ai server Microsoft."); }
}

function aggiornaInterfaccia() {
    const azione = document.getElementById("actionSelector").value;
    const input = document.getElementById("dynamicInput");
    const btn = document.getElementById("executeBtn");
    
    input.style.display = "block";
    btn.classList.remove("danger");
    triggerAnimazione();

    const azioniSenzaTesto = ["markRead", "flag", "pdf", "open", "delete", "security"];
    
    if (azioniSenzaTesto.includes(azione)) {
        input.style.display = "none";
        if (azione === "delete") btn.classList.add("danger");
    } else if (azione === "search") input.placeholder = "Scrivi la parola da cercare...";
    else if (azione === "folder") input.placeholder = "Nome della cartella...";
    else if (azione === "tag") input.placeholder = "Nome del Tag...";
    else if (azione === "spam") input.placeholder = "Indirizzo email da bloccare...";
    
    input.value = "";
}

function eseguiAzioneScelta() {
    const azione = document.getElementById("actionSelector").value;
    
    if (azione === "search") cercaEmailIntelligente();
    else if (azione === "folder") spostaInCartellaMassa();
    else if (azione === "tag") applicaTagMassa();
    else if (azione === "spam") bloccaMittente();
    else if (azione === "delete") eliminaMassa();
    else if (azione === "markRead") segnaLette();
    else if (azione === "flag") contrassegna();
    else if (azione === "pdf") salvaInPDF();
    else if (azione === "open") apriNelBrowser();
    else if (azione === "security") analizzaSicurezza();
}

// RICERCA AMPLIATA CON I LINK WEB
async function cercaEmailIntelligente() {
    if (!accessToken) return;
    const query = document.getElementById("dynamicInput").value.trim();
    if (!query) return alert("Scrivi qualcosa da cercare!");
    
    aggiornaStato("🔍 Ricerca in corso...");
    document.getElementById("empty-state").style.display = "none";
    
    // NOTA: Aggiunto webLink per permettere l'apertura nel browser
    const urlGraph = `https://graph.microsoft.com/v1.0/me/messages?$search="${query}"&$top=50&$select=id,subject,from,webLink`;

    try {
        const response = await fetch(urlGraph, { headers: { 'Authorization': `Bearer ${accessToken}` } });
        const data = await response.json();
        foundEmails = data.value || [];
        
        mostraRisultati();
        aggiornaStato(`Trovate ${foundEmails.length} email.`);
        document.getElementById("selectAllBtn").classList.remove("hidden");
    } catch (error) { aggiornaStato("⚠️ Errore di ricerca."); }
}

function mostraRisultati() {
    const area = document.getElementById("results-area");
    area.innerHTML = ''; 
    
    if (foundEmails.length === 0) {
        area.innerHTML = `<div style="text-align: center; color: var(--text-sec); font-size: 12px; margin-top: 20px;">Nessun risultato trovato.</div>`;
        document.getElementById("selectAllBtn").classList.add("hidden");
        return;
    }
    
    foundEmails.forEach(email => {
        const mittente = email.from ? email.from.emailAddress.name : "Sconosciuto";
        const emailAddress = email.from ? email.from.emailAddress.address : "";
        const div = document.createElement("div");
        div.className = "email-item animate-ui"; 
        div.innerHTML = `
            <input type="checkbox" class="mail-checkbox" value="${email.id}" data-email="${emailAddress}" data-link="${email.webLink}">
            <div class="email-info">
                <div class="email-sender">${mittente} <span style="font-size:10px; font-weight:normal;">${emailAddress}</span></div>
                <div class="email-subject">${email.subject || '(Senza Oggetto)'}</div>
            </div>
        `;
        area.appendChild(div);
    });
}

function toggleSelezionaTutte() {
    const checkboxes = document.querySelectorAll(".mail-checkbox");
    const tuttiSelezionati = Array.from(checkboxes).every(cb => cb.checked);
    checkboxes.forEach(cb => cb.checked = !tuttiSelezionati);
}
function getIdSelezionati() { return Array.from(document.querySelectorAll(".mail-checkbox:checked")).map(cb => cb.value); }

// --- NUOVE FUNZIONALITÀ ---

// 1. ANALISI SICUREZZA (Euristica Base)
function analizzaSicurezza() {
    const checkboxes = document.querySelectorAll(".mail-checkbox:checked");
    if (checkboxes.length !== 1) return alert("Seleziona esattamente UNA mail per analizzarla.");
    
    const emailIndirizzo = checkboxes[0].getAttribute("data-email");
    aggiornaStato(`🛡️ Analisi di: ${emailIndirizzo}...`);
    
    // Controlli euristici basilari
    let punteggioRischio = 0;
    const dominiFidati = ['microsoft.com', 'google.com', 'apple.com', 'amazon.it', 'paypal.com'];
    const dominio = emailIndirizzo.split('@')[1] || "";
    
    if (emailIndirizzo.match(/[0-9]{4,}/)) punteggioRischio += 50; // Troppi numeri
    if (emailIndirizzo.includes("noreply") || emailIndirizzo.includes("update")) punteggioRischio += 20;
    if (!dominiFidati.includes(dominio.toLowerCase())) punteggioRischio += 30; // Dominio non noto

    setTimeout(() => {
        if (punteggioRischio >= 70) {
            alert(`⚠️ RISCHIO ALTO: L'indirizzo ${emailIndirizzo} ha caratteristiche sospette (Numeri strani o dominio non verificato). Ti consiglio di usare la funzione 'Blocca Mittente (Spam)'.`);
            aggiornaStato("⚠️ Mittente sospetto rilevato.");
        } else {
            alert(`✅ RISCHIO BASSO: L'indirizzo ${emailIndirizzo} sembra avere una struttura normale. Cautela comunque coi link!`);
            aggiornaStato("✅ Sembra sicuro.");
        }
    }, 1000);
}

// 2. SEGNA COME LETTE
async function segnaLette() {
    const ids = getIdSelezionati();
    if (ids.length === 0) return alert("Seleziona le mail!");
    aggiornaStato(`⏳ Segno ${ids.length} email come lette...`);
    
    for (const id of ids) {
        try {
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
                method: 'PATCH',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ isRead: true })
            });
        } catch(e) {}
    }
    aggiornaStato("🟢 Email segnate come lette!");
}

// 3. CONTRASSEGNA (Bandierina)
async function contrassegna() {
    const ids = getIdSelezionati();
    if (ids.length === 0) return alert("Seleziona le mail!");
    aggiornaStato(`⏳ Contrassegno ${ids.length} email...`);
    
    for (const id of ids) {
        try {
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
                method: 'PATCH',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ flag: { flagStatus: 'flagged' } })
            });
        } catch(e) {}
    }
    aggiornaStato("🚩 Email contrassegnate!");
}

// 4. APRI NEL BROWSER (Nuova Finestra)
function apriNelBrowser() {
    const checkboxes = document.querySelectorAll(".mail-checkbox:checked");
    if (checkboxes.length === 0) return alert("Seleziona le mail da aprire!");
    
    checkboxes.forEach(cb => {
        const link = cb.getAttribute("data-link");
        if(link) window.open(link, '_blank');
    });
    aggiornaStato("🌐 Finestre aperte nel browser.");
}

// 5. STAMPA / SALVA COME PDF (Trucco del Programmatore)
async function salvaInPDF() {
    const ids = getIdSelezionati();
    if (ids.length !== 1) return alert("Seleziona UNA SOLA mail per salvarla in PDF.");
    
    aggiornaStato(`⏳ Generazione PDF in corso...`);
    
    try {
        // Scarichiamo il corpo completo dell'email
        const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${ids[0]}?$select=subject,body`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        const data = await res.json();
        
        // Creiamo una finestra temporanea per la stampa nativa
        const printWindow = window.open('', '_blank');
        printWindow.document.write(`
            <html><head><title>${data.subject || 'Email PDF'}</title></head>
            <body style="font-family:sans-serif; padding:20px;">
                <h2>${data.subject || 'Senza Oggetto'}</h2>
                <hr/>
                ${data.body.content}
            </body></html>
        `);
        printWindow.document.close();
        // Chiamiamo la funzione di stampa del PC
        printWindow.focus();
        setTimeout(() => { printWindow.print(); }, 1000);
        
        aggiornaStato("📄 Scegli 'Salva come PDF' dalla finestra di stampa!");
    } catch(e) {
        aggiornaStato("⚠️ Errore durante la creazione del PDF.");
    }
}

// VECCHIE FUNZIONI (Cartelle, Tag, Spam, Elimina)
async function spostaInCartellaMassa() { /* Omissis: identica a prima, assicurati di mantenerla! */ }
async function applicaTagMassa() { /* Omissis: identica a prima */ }
async function bloccaMittente() { /* Omissis: identica a prima */ }
async function eliminaMassa() { /* Omissis: identica a prima */ }

// --- [INCOLLA QUI IL RESTO DELLE FUNZIONI VECCHIE (Sposta, Tagga, Elimina, Spam) CHE GIA' AVEVAMO] ---
