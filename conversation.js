(function (global) {
    "use strict";

    function create(options = {}) {
        let messages = [], busy = false, counter = 0;
        let saving = Promise.resolve();
        const clean = items => (Array.isArray(items) ? items : []).slice(-40)
            .filter(item => item && ["user", "assistant"].includes(item.role) && typeof item.content === "string")
            .map(item => ({role:item.role, content:item.content.slice(0, 6000),
                kind:["request", "done", "question", "answer", "error"].includes(item.kind) ? item.kind : "answer",
                tableId:typeof item.tableId === "string" ? item.tableId : null}));
        const changed = () => options.onChange?.(messages.map(item => ({...item})), busy);
        function persist() {
            const copy = clean(messages);
            saving = saving.catch(() => {}).then(() => options.save?.({version:1, messages:copy}))
                .catch(() => options.onStorageError?.());
        }
        const ready = Promise.resolve().then(() => options.load?.()).then(saved => {
            if (saved?.version === 1) messages = clean(saved.messages);
        }).catch(() => options.onStorageError?.()).then(changed);
        return {
            ready,
            add(role, content, kind = "answer", tableId = null) {
                const id = ++counter;
                messages = clean([...messages, {role, content, kind, tableId}]);
                changed(); persist(); return id;
            },
            history() {
                return messages.slice(-12).map(item => ({role:item.role, content:item.content.slice(0, 3000)}));
            },
            lastQuestion() {
                const last = [...messages].reverse().find(item => item.role === "assistant");
                return last?.kind === "question" ? {...last} : null;
            },
            setBusy(value) {busy = Boolean(value); changed();},
            setContext(value) {options.onContext?.(String(value || ""));},
            reset() {
                if (busy) return false;
                messages = []; changed(); persist(); options.onReset?.(); return true;
            },
            snapshot() {return messages.map(item => ({...item}));},
            flushed() {return saving;}
        };
    }

    function mount(options) {
        const doc = options.prompt.ownerDocument || global.document;
        if (!options.prompt.parentElement) return create(options);
        const style = doc.createElement("style");
        style.textContent = `
.tableia-chat{margin:12px 0 16px;text-align:left;font-family:inherit;color:#243345}
.tableia-chat-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:8px}
.tableia-chat-title{font-size:17px;font-weight:750;margin:0}
.tableia-chat-reset{border:1px solid #ccd8d1;background:#fff;border-radius:9px;padding:7px 9px;color:#176940;font:inherit;font-size:12px;cursor:pointer}
.tableia-chat-reset:disabled{opacity:.55;cursor:wait}
.tableia-chat-context{font-size:12px;color:#52685c;background:#edf6f0;border-radius:8px;padding:8px 10px;margin-bottom:10px;overflow-wrap:anywhere}
.tableia-chat-log{max-height:340px;overflow-y:auto;overscroll-behavior:contain;padding:2px 4px;scrollbar-width:thin}
.tableia-chat-message{max-width:94%;margin:10px 0;padding:10px 12px;border-radius:12px;background:#f1f4f7;box-sizing:border-box}
.tableia-chat-message[data-role="user"]{margin-left:auto;background:#e5f3ea}
.tableia-chat-message[data-kind="error"]{background:#fff0ed}
.tableia-chat-author{font-size:11px;font-weight:700;color:#52645d;display:block;margin-bottom:5px}
.tableia-chat-text{font-size:14px;line-height:1.5;white-space:pre-wrap;overflow-wrap:anywhere;margin:0}
.tableia-chat-empty{font-size:14px;line-height:1.5;color:#607168;padding:10px 2px}
.tableia-chat-hint{font-size:11px;color:#718078;margin-top:8px}
.tableia-chat-input{box-sizing:border-box;min-height:88px!important;height:96px;resize:vertical}
@media(max-width:360px){.tableia-chat-head{align-items:flex-start}.tableia-chat-log{max-height:280px}}
`;
        (doc.head || doc.body).appendChild(style);
        const root = doc.createElement("section"); root.className = "tableia-chat";
        root.setAttribute("aria-label", "Conversation avec TableIA");
        const head = doc.createElement("div"); head.className = "tableia-chat-head";
        const title = doc.createElement("h2"); title.className = "tableia-chat-title"; title.textContent = "Votre conversation";
        const reset = doc.createElement("button"); reset.type = "button"; reset.className = "tableia-chat-reset"; reset.textContent = "Nouvelle conversation";
        head.appendChild(title); head.appendChild(reset);
        const context = doc.createElement("div"); context.className = "tableia-chat-context";
        const log = doc.createElement("div"); log.className = "tableia-chat-log";
        log.setAttribute("role", "log"); log.setAttribute("aria-live", "polite"); log.setAttribute("aria-relevant", "additions text");
        const hint = doc.createElement("div"); hint.className = "tableia-chat-hint";
        hint.textContent = "Entrée pour envoyer · Maj + Entrée pour une nouvelle ligne";
        root.appendChild(head); root.appendChild(context); root.appendChild(log); root.appendChild(hint);
        options.prompt.parentElement.insertBefore(root, options.prompt);
        options.prompt.classList.add("tableia-chat-input");
        options.prompt.placeholder = "Écrivez votre demande ou votre réponse…";
        options.prompt.setAttribute("aria-label", "Message à TableIA");
        const controller = create({...options,
            onChange(messages, busy) {
                log.replaceChildren(); reset.disabled = busy;
                root.setAttribute("aria-busy", String(busy));
                if (!messages.length) {
                    const empty = doc.createElement("p"); empty.className = "tableia-chat-empty";
                    empty.textContent = "Créez un tableau, précisez une demande ou demandez de l’aide. Je m’appuie sur le tableau Excel actif.";
                    log.appendChild(empty);
                }
                for (const message of messages) {
                    const bubble = doc.createElement("div"); bubble.className = "tableia-chat-message";
                    bubble.dataset.role = message.role; bubble.dataset.kind = message.kind;
                    const author = doc.createElement("span"); author.className = "tableia-chat-author";
                    author.textContent = message.role === "user" ? "Vous" : message.kind === "done" ? "TableIA · Action effectuée" : "TableIA";
                    const content = doc.createElement("p"); content.className = "tableia-chat-text"; content.textContent = message.content;
                    bubble.appendChild(author); bubble.appendChild(content); log.appendChild(bubble);
                }
                log.scrollTop = log.scrollHeight;
                options.onChange?.(messages, busy);
            },
            onContext(value) {context.textContent = value;},
            onStorageError() {hint.textContent = "Historique disponible pour cette session · Entrée pour envoyer";}
        });
        reset.addEventListener("click", () => controller.reset());
        options.prompt.addEventListener("keydown", event => {
            if (event.key === "Enter" && !event.shiftKey && !event.isComposing) {
                event.preventDefault(); options.onSend?.();
            }
        });
        return controller;
    }

    global.TableIAConversation = {create, mount};
})(window);
