import win32com.client


# Only variable value that you should change
usernames = ["Firstname Lastname", "Firstname1 Lastname1", "Firstname2 Lastname2", "Firstname3 Lastname3"]
# -----


outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
global_address_list = outlook.Session.GetGlobalAddressList()
address_book_entries = global_address_list.AddressEntries

final_result = ""

for i in range(len(usernames)):
    try:
        identifier = address_book_entries[usernames[i]]

        email_address = None
        if "EX" == identifier.Type:     # Double if statement intended because you want the email_address variable to be overwritten if necessary
            eu = identifier.GetExchangeUser()
            email_address = eu.PrimarySmtpAddress

        if "SMTP" == identifier.Type:
            email_address = identifier.Address

        final_result += email_address + "; "
    except:
        pass

print(f"\nEmail addresses are:\n\n{final_result}\n")