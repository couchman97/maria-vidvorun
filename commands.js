Office.onReady();

function onMessageSendHandler(event) {
    // ⚠️ LISTINN yfir netföng sem á að vara við — bætt við eftir þörfum
    const warnAddresses = [
        "maria@ronning.is"  // ← BREYTTU í raunverulegt netfang
    ];
    
    const item = Office.context.mailbox.item;
    
    Promise.all([
        new Promise(r => item.to.getAsync(res => r(res.value || []))),
        new Promise(r => item.cc.getAsync(res => r(res.value || [])))
    ]).then(([to, cc]) => {
        const allRecipients = [...to, ...cc];
        const hit = allRecipients.find(r => 
            r.emailAddress && warnAddresses.some(w => 
                r.emailAddress.toLowerCase() === w.toLowerCase()
            )
        );
        
        if (hit) {
            event.completed({
                allowEvent: false,
                cancelMessage: `⚠️ STOPP! Þú ert að senda á ${hit.displayName} (${hit.emailAddress}). Þetta er María hjá Rönning — ekki María Rós hjá RST. Fjarlægðu hana og reyndu aftur ef þetta er ekki rétt.`
            });
        } else {
            event.completed({ allowEvent: true });
        }
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);