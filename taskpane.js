// Stub per la retrocompatibilità con i vecchi sistemi Office
Office.initialize = function () {};

let myMSALObj;
let accessToken = null;
let foundEmails = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        
        const msalConfig = {
            auth: {
                clientId: "65c3d535-927c-4930-b0e9-0cefddf82495", 
                authority: "https://login.microsoftonline.com/common",
                redirectUri: "https://gaetano-taormina.github.io/outlook-chatbot/taskpane.html"
            },
            cache: {
                cacheLocation: "localStorage",
                storeAuthStateInCookie: true
            }
        };
        
        myMSALObj = new msal.PublicClientApplication(msalConfig);

        // Collegamento eventi Login
        document.getElementById("loginBtn").addEventListener("click", login);
        
        // Logica dell'Interfaccia Dinamica
        document.getElementById("actionSelector").addEventListener("change", aggiornaInterfaccia);
        document.getElementById("executeBtn").addEventListener("click", eseguiAzioneScelta);
        document.getElementById("selectAllBtn").addEventListener("click", toggleSelezionaTutte);
    }
});

function aggiornaStato(testo) {
    const statusEl = document.getElementById("status-text");
    if (statusEl) statusEl.innerText = testo;
}

async function login() {
    try {
        const loginRequest = { scopes: ["user.read", "mail.readwrite"] };
        const loginResponse = await myMSALObj.loginPopup(loginRequest);
        
        const tokenRequest = { scopes: ["mail.readwrite"], account: loginResponse.account };
        const tokenResponse = await myMSALObj.acquireTokenSilent(tokenRequest);
        accessToken = tokenResponse.accessToken;
        
        // Transizione dalla Schermata Login all'App
        document.getElementById("login-screen").style.display = "none";
        document.getElementById("app-screen").style.display = "flex";
        aggiornaStato("🟢 Connesso. Seleziona un'azione e scrivi nel box.");
        
    } catch (error) {
        console.error(error);
        alert("Errore di connessione. Controlla i popup del browser.");
    }
}

// LOGICA CAMBIO UI (Il menu a tendina)
function aggiornaInterfaccia() {
    const azione = document.getElementById("actionSelector").value;
    const input = document.getElementById("dynamicInput");
    const btn = document.getElementById("executeBtn");
    
    input.style.display = "block";
    btn.classList.remove("danger");

    if (azione === "search") {
        input.placeholder = "Scrivi la parola da cercare...";
        input.value = "";
    } else if (azione === "folder") {
        input.placeholder = "Nome della cartella di destinazione...";
        input.value = "";
    } else if (azione === "tag") {
        input.placeholder = "Nome del Tag da applicare...";
        input.value = "";
    } else if (azione === "delete") {
        input.style.display = "none"; // Non serve scrivere nulla per eliminare
        btn.classList.add("danger");
    }
}

// SMISTATORE DEL PULSANTE "INVIO"
function eseguiAzioneScelta() {
    const azione = document.getElementById("actionSelector").value;
    
    if (azione === "search") cercaEmailIntelligente();
    else if (azione === "folder") spostaInCartellaMassa();
    else if (azione === "tag") applicaTagMassa();
    else if (azione === "delete") eliminaMassa();
}

async function cercaEmailIntelligente() {
    if (!accessToken) return;
    
    const query = document.getElementById("dynamicInput").value.trim();
    if (!query) return alert("Scrivi qualcosa da cercare!");
    
    aggiornaStato("🔍 Ricerca in corso...");
    document.getElementById("empty-state").style.display = "none";
    
    // Ricerca super flessibile: non distingue maiuscole/minuscole
    const urlGraph = `https://graph.microsoft.com/v1.0/me/messages?$search="${query}"&$top=50&$select=id,subject,from`;

    try {
        const response = await fetch(urlGraph, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        const data = await response.json();
        foundEmails = data.value || [];
        
        mostraRisultati();
        aggiornaStato(`Trovate ${foundEmails.length} email.`);
        document.getElementById("selectAllBtn").classList.remove("hidden");
    } catch (error) {
        aggiornaStato("⚠️ Errore di ricerca.");
    }
}

function mostraRisultati() {
    const area = document.getElementById("results-area");
    // Puliamo l'area (lasciando l'empty state nascosto)
    area.innerHTML = ''; 
    
    if (foundEmails.length === 0) {
        area.innerHTML = `<div style="text-align: center; color: var(--text-sec); font-size: 12px; margin-top: 20px;">Nessun risultato trovato.</div>`;
        document.getElementById("selectAllBtn").classList.add("hidden");
        return;
    }
    
    foundEmails.forEach(email => {
        const mittente = email.from ? email.from.emailAddress.name : "Sconosciuto";
        const div = document.createElement("div");
        div.className = "email-item";
        div.innerHTML = `
            <input type="checkbox" class="mail-checkbox" value="${email.id}">
            <div class="email-info">
                <div class="email-sender">${mittente}</div>
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

function getIdSelezionati() {
    return Array.from(document.querySelectorAll(".mail-checkbox:checked")).map(cb => cb.value);
}

async function spostaInCartellaMassa() {
    const ids = getIdSelezionati();
    const nomeCartella = document.getElementById("dynamicInput").value.trim();
    
    if (ids.length === 0) return alert("Seleziona almeno una mail sopra!");
    if (!nomeCartella) return alert("Scrivi il nome della Cartella!");
    
    aggiornaStato(`⏳ Creazione cartella e spostamento...`);
    
    try {
        let folderId;
        const resCheck = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders?$filter=displayName eq '${nomeCartella}'`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        const dataCheck = await resCheck.json();
        
        if (dataCheck.value && dataCheck.value.length > 0) folderId = dataCheck.value[0].id;
        else {
            const resCreate = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders`, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ displayName: nomeCartella, isHidden: false })
            });
            const dataCreate = await resCreate.json();
            folderId = dataCreate.id;
        }

        for (const id of ids) {
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}/move`, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ destinationId: folderId })
            });
        }
        
        aggiornaStato(`🟢 Completato! Situate in "${nomeCartella}".`);
        cercaEmailIntelligente(); 
    } catch (e) { aggiornaStato("⚠️ Errore durante lo spostamento."); }
}

async function applicaTagMassa() {
    const ids = getIdSelezionati();
    const nuovoTag = document.getElementById("dynamicInput").value.trim();
    
    if (ids.length === 0) return alert("Seleziona almeno una mail!");
    if (!nuovoTag) return alert("Scrivi il nome del Tag!");
    
    aggiornaStato(`⏳ Sto colorando le email...`);
    
    for (const id of ids) {
        try {
            const mailOriginale = foundEmails.find(m => m.id === id);
            let cat = mailOriginale.categories || [];
            if (!cat.includes(nuovoTag)) cat.push(nuovoTag);
            
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
                method: 'PATCH',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ categories: cat })
            });
        } catch (e) { }
    }
    aggiornaStato("🟢 Tag applicati con successo!");
}

async function eliminaMassa() {
    const ids = getIdSelezionati();
    if (ids.length === 0) return alert("Seleziona le mail da eliminare!");
    
    if (!confirm(`Sei sicuro di voler eliminare ${ids.length} email?`)) return;
    
    aggiornaStato(`⏳ Sto cestinando le email...`);
    for (const id of ids) {
        try {
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
                method: 'DELETE',
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
        } catch (e) { }
    }
    aggiornaStato("🟢 Pulizia completata!");
    cercaEmailIntelligente(); 
}
