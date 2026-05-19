// Inizializzazione obbligatoria per Outlook
Office.initialize = function () {};

Office.onReady((info) => {
    // Outlook ora sa che il bot è vivo, quindi permette di mostrare la grafica
    if (info.host === Office.HostType.Outlook) {
        
        // Riferimenti DOM
        const chatArea = document.getElementById("chat-area");
        const typingIndicator = document.getElementById("typingIndicator");
        const wifiIcon = document.getElementById("wifiIcon");
        const menuBtn = document.getElementById("menuToggleBtn");
        const inputField = document.getElementById("dynamicInput");
        const resetBtn = document.getElementById("resetBtn");
        
        // Funzione Utility: Inserisci messaggio
        function aggiungiMessaggio(testo, tipo) {
            const div = document.createElement("div");
            div.className = tipo === 'bot' ? 'bot-msg' : 'user-msg';
            div.innerHTML = testo;
            chatArea.insertBefore(div, typingIndicator);
            chatArea.scrollTop = chatArea.scrollHeight; // Scroll automatico
        }

        // Funzione Utility: Mostra/Nascondi typing
        function setTyping(mostra) {
            typingIndicator.style.display = mostra ? "flex" : "none";
            if(mostra) chatArea.scrollTop = chatArea.scrollHeight;
        }

        // ----------------------------------------------------
        // SEQUENZA DI AVVIO SIMULATA (Finto Login)
        // ----------------------------------------------------
        setTimeout(() => {
            // Il Wi-Fi diventa giallo: Sto connettendo...
            wifiIcon.className = "wifi-icon wifi-yellow";
            
            setTimeout(() => {
                // Wi-Fi Verde: Connesso!
                wifiIcon.className = "wifi-icon wifi-green";
                inputField.disabled = false;
                inputField.placeholder = "Scrivi la tua richiesta...";
                
                // Il bot scrive...
                setTyping(true);
                
                setTimeout(() => {
                    setTyping(false);
                    aggiungiMessaggio("Ciao! Con quale azione vuoi cominciare oggi?<br><br><b>Opzioni suggerite:</b><br>• Cerca email<br>• Cerca mittenti<br>• Esegui comandi / Regole", "bot");
                    
                    // Il tasto menu inizia a pulsare
                    menuBtn.classList.add("menu-pulsing");

                }, 1500);

            }, 2000); // Finto caricamento connessione

        }, 500);

        // ----------------------------------------------------
        // LOGICA TASTO RESET (Aggiorna)
        // ----------------------------------------------------
        resetBtn.addEventListener("click", () => {
            // 1. Animazione freccia fulminea
            resetBtn.classList.remove("spin-anim");
            void resetBtn.offsetWidth; // Forza reflow
            resetBtn.classList.add("spin-anim");

            // 2. Cancella tutta la chat (tranne l'indicatore typing)
            const messages = chatArea.querySelectorAll('.bot-msg, .user-msg');
            messages.forEach(msg => msg.remove());
            
            // Ferma il menu pulsante
            menuBtn.classList.remove("menu-pulsing");
            
            // 3. Riparte l'animazione di scrittura
            setTyping(true);
            
            // 4. Dopo ~2 secondi riappare il benvenuto e il menu pulsa
            setTimeout(() => {
                setTyping(false);
                aggiungiMessaggio("Ciao! Ripartiamo da zero. Cosa vuoi fare?<br><br><b>Opzioni suggerite:</b><br>• Cerca email<br>• Cerca mittenti<br>• Esegui comandi / Regole", "bot");
                menuBtn.classList.add("menu-pulsing");
            }, 2000);
        });

        // Ferma la pulsazione se l'utente clicca il menu
        menuBtn.addEventListener("click", () => {
            menuBtn.classList.remove("menu-pulsing");
            aggiungiMessaggio("Hai cliccato il menu. Qui appariranno le tue opzioni.", "bot");
        });
    }
});
