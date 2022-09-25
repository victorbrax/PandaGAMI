from win10toast import ToastNotifier 

toaster = ToastNotifier()


def success_notification(qtemails):
    toaster.show_toast(
        "We won again, Sencineer!", 
        f"Foram enviados {qtemails} e-mails em nome da Queries FSSC! ;)", 
        icon_path='utils/notification/iconcheck.ico', 
        duration=5, 
        threaded=False, 
        )


def start_working():
    toaster.show_toast(
        "Iniciando os trabalhos", 
        "Já pegamos os dados, hora de enviar as mensagens...", 
        icon_path='utils/notification/dataconfirm.ico', 
        duration=5, 
        threaded=False, 
        )
