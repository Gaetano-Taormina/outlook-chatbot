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

        document.getElementById("loginBtn").addEventListener("click", login);
        document.getElementById("searchBtn").addEventListener("click", cercaEmailIntelligente);
        document.getElementById("tagBtn").addEventListener("click", applicaTagMassa);
        document.getElementById("folderBtn").addEventListener("click", spostaInCartellaMassa);
        document.getElementById("deleteBtn").addEventListener("click", eliminaMassa);
        document.getElementById("selectAllBtn").addEventListener("click", toggleSelezionaTutte);
    }
});

function aggiornaStato(testo, colore = "#333") {
    const statusEl = document.getElementById("status");
    if (statusEl) {
        statusEl.innerText = testo;
        statusEl.style.color = colore;
    }
}

async function login() {
    try {
        aggiornaStato("⏳ Accesso in corso...", "#d83b01");
        const loginRequest = { scopes: ["user.read", "mail.readwrite"] };
        const loginResponse = await myMSALObj.loginPopup(loginRequest);
        
        const tokenRequest = { scopes: ["mail.readwrite"], account: loginResponse.account };
        const tokenResponse = await myMSALObj.acquireTokenSilent(tokenRequest);
        accessToken = tokenResponse.accessToken;
        
        aggiornaStato("🟢 Connesso e Pronto all'uso!", "#107c41");
        document.getElementById("loginBtn").style.display = "none"; 
    } catch (error) {
        console.error(error);
        aggiornaStato("⚠️ Errore di connessione.", "#a80000");
    }
}

async function cercaEmailIntelligente() {
    if (!accessToken) { return alert("Devi prima connettere l'account!"); }
    
    const includi = document.getElementById("includeQuery").value.trim();
    const escludi = document.getElementById("excludeQuery").value.trim();
    
    if (!includi) { return alert("Scrivi almeno una parola da cercare nel campo 'Trova'!"); }
    
    aggiornaStato("🔍 Sto cercando le mail...", "#0078d4");
    
    let logicaDiRicerca = `"${includi}"`;
    if (escludi) { logicaDiRicerca += ` NOT "${escludi}"`; }

    const urlGraph = `https://graph.microsoft.com/v1.0/me/messages?$search=${logicaDiRicerca}&$top=50&$select=id,subject,from,categories`;

    try {
        const response = await fetch(urlGraph, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        const data = await response.json();
        foundEmails = data.value || [];
        
        mostraRisultati();
        aggiornaStato("🟢 Ricerca completata!", "#107c41");
    } catch (error) {
        aggiornaStato("⚠️ Errore di ricerca.", "#a80000");
    }
}

function mostraRisultati() {
    const listaDiv = document.getElementById("emailList");
    listaDiv.innerHTML = "";
    document.getElementById("matchCount").innerText = foundEmails.length;
    
    if (foundEmails.length === 0) {
        listaDiv.innerHTML = "<p style='padding:10px; font-size:12px; text-align:center;'>Nessuna mail trovata con questi filtri.</p>";
        document.getElementById("resultsSection").style.display = "block";
        return;
    }
    
    foundEmails.forEach(email => {
        const mittente = email.from ? email.from.emailAddress.name : "Sconosciuto";
        const div = document.createElement("div");
        div.className = "email-item";
        div.innerHTML = `
            <input type="checkbox" class="mail-checkbox" value="${email.id}">
            <div style="margin-left: 8px;"><strong>${mittente}</strong><br><span style="color:#888">${email.subject || '(Senza Oggetto)'}</span></div>
        `;
        listaDiv.appendChild(div);
    });
    
    document.getElementById("resultsSection").style.display = "block";
}

function toggleSelezionaTutte() {
    const checkboxes = document.querySelectorAll(".mail-checkbox");
    const tuttiSelezionati = Array.from(checkboxes).every(cb => cb.checked);
    checkboxes.forEach(cb => cb.checked = !tuttiSelezionati);
}

function getIdSelezionati() {
    return Array.from(document.querySelectorAll(".mail-checkbox:checked")).map(cb => cb.value);
}

// IL NUOVO MOTORE PER LE CARTELLE STILE GMAIL
async function spostaInCartellaMassa() {
    const ids = getIdSelezionati();
    const nomeCartella = document.getElementById("newTag").value.trim();
    
    if (ids.length === 0) return alert("Spunta almeno una mail dalla lista!");
    if (!nomeCartella) return alert("Scrivi il nome della Cartella da usare!");
    
    aggiornaStato(`⏳ Creazione cartella e spostamento di ${ids.length} email...`, "#d83b01");
    
    try {
        // 1. Controlliamo se la cartella esiste già
        let folderId;
        const resCheck = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders?$filter=displayName eq '${nomeCartella}'`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        const dataCheck = await resCheck.json();
        
        if (dataCheck.value && dataCheck.value.length > 0) {
            folderId = dataCheck.value[0].id; // Esiste già
        } else {
            // 2. Non esiste, la creiamo
            const resCreate = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders`, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ displayName: nomeCartella, isHidden: false })
            });
            const dataCreate = await resCreate.json();
            folderId = dataCreate.id;
        }

        // 3. Spostiamo le email selezionate nella cartella
        for (const id of ids) {
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}/move`, {
                method: 'POST',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ destinationId: folderId })
            });
        }
        
        aggiornaStato(`🟢 Completato! Situate in "${nomeCartella}".`, "#107c41");
        cercaEmailIntelligente(); // Aggiorna la lista
        
    } catch (e) {
        console.error(e);
        aggiornaStato("⚠️ Errore durante lo spostamento.", "#a80000");
    }
}

async function applicaTagMassa() {
    const ids = getIdSelezionati();
    const nuovoTag = document.getElementById("newTag").value.trim();
    
    if (ids.length === 0) return alert("Spunta almeno una mail dalla lista!");
    if (!nuovoTag) return alert("Scrivi il nome del Tag che vuoi applicare!");
    
    aggiornaStato(`⏳ Sto colorando ${ids.length} email...`, "#d83b01");
    
    for (const id of ids) {
        try {
            const mailOriginale = foundEmails.find(m => m.id === id);
            let categorieAttuali = mailOriginale.categories || [];
            if (!categorieAttuali.includes(nuovoTag)) categorieAttuali.push(nuovoTag);
            
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
                method: 'PATCH',
                headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
                body: JSON.stringify({ categories: categorieAttuali })
            });
        } catch (e) { console.error(e); }
    }
    
    aggiornaStato("🟢 Tag applicati con successo!", "#107c41");
    cercaEmailIntelligente(); 
}

async function eliminaMassa() {
    const ids = getIdSelezionati();
    if (ids.length === 0) return alert("Spunta le mail che vuoi eliminare!");
    
    if (!confirm(`Sei assolutamente sicuro di voler cestinare ${ids.length} email?`)) return;
    
    aggiornaStato(`⏳ Sto cestinando ${ids.length} email...`, "#a80000");
    
    for (const id of ids) {
        try {
            await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
                method: 'DELETE',
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
        } catch (e) { console.error(e); }
    }
    
    aggiornaStato("🟢 Pulizia completata!", "#107c41");
    cercaEmailIntelligente(); 
}
