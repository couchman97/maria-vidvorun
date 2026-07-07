Office.onReady();

function onMessageSendHandler(event) {
    const warnAddresses = [
        "maria@ronning.is"
    ];
    
    // Oryggis-timeout - leyfa sendingu ef eitthvad hangir
    const timeout = setTimeout(() => {
        event.completed({ allowEvent: true });
    }, 2000);
    
    const item = Office.context.mailbox.item;
    
    item.to.getAsync(toResult => {
        item.cc.getAsync(ccResult => {
            clearTimeout(timeout);
            
            const all = (toResult.value || []).concat(ccResult.value || []);
            
            let match = null;
            for (let i = 0; i < all.length; i++) {
                const email = (all[i].emailAddress || "").toLowerCase().trim();
                for (let j = 0; j < warnAddresses.length; j++) {
                    if (email === warnAddresses[j].toLowerCase().trim()) {
                        match = all[i];
                        break;
                    }
                }
                if (match) break;
            }
            
            if (match) {
                event.completed({
                    allowEvent: false,
                    cancelMessage: "STOPP! Thu ert ad senda a " + match.displayName + " (" + match.emailAddress + ")."
                });
            } else {
                event.completed({ allowEvent: true });
            }
        });
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);