def direcionarDownloads(id):
    #permitir notificaçoes e download automatico
    global prefs
    prefs = {"profile.default_content_setting_values.notifications" : 1, "profile.default_content_setting_values.automatic_downloads": 1,
            "download.default_directory": r"C:\\Users\\Usuario\\OneDrive - tpfe.com.br\\" + str(data_em_texto) + "\\"+ str(id)}

def opçoesdoChrome():
    return prefs