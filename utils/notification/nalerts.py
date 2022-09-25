from win10toast import ToastNotifier 

toaster = ToastNotifier()


def success_notification(qtemails):
    toaster.show_toast(
        "¡Ganamos de nuevo, Sencineer!", 
        f"¡Se enviaron {qtemails} correos electrónicos en nombre de Queries FSSC! ;)", 
        icon_path='utils/notification/iconcheck.ico', 
        duration=5, 
        threaded=False, 
        )


def start_working():
    toaster.show_toast(
        "Empezando a trabajar...", 
        "Ya tenemos los datos, hora de enviar los mensajes...", 
        icon_path='utils/notification/dataconfirm.ico', 
        duration=5, 
        threaded=False, 
        )
