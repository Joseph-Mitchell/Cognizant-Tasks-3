async function saveContactIfRequired(executionContext) {
    var formContext = executionContext.getFormContext();

    if (!executionContext.getSharedVariable("showAccountPrimaryContactFlags").saveNewPrimaryContact)
        return;
    
    var accountId = formContext.getAttribute("customerid").getValue()[0].id;
    var contactId = formContext.getAttribute("primarycontactid").getValue()[0].id;
    await saveAccountPrimaryContact(accountId, contactId);
    
    cleanup();
}

async function saveAccountPrimaryContact(accountId, contactId) {
    try {
        await Xrm.WebApi.updateRecord(
            "account",
            accountId,
            {
                "primarycontactid@odata.bind": `/contacts(${contactId})`
            }
        );
    } catch {
        Xrm.Navigation.openErrorDialog({ message: "Error saving primary contact to account" });
    }
}

function cleanup(executionContext, formContext) {
    executionContext.getSharedVariable("showAccountPrimaryContactFlags", { saveNewPrimaryContact: false });
    
    var contactAttribute = formContext.getAttribute("primarycontactid");
    var contactControl = contactAttribute.controls.get((control) => {
        if (control.controlDescriptor.Label == "Account Primary Contact")
            return true;
    })[0];
    
    contactAttribute.setRequiredLevel("none");
    contactControl.setVisible(false);
}