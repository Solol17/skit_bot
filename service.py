import socket
import win32serviceutil
import win32service
import win32event
import servicemanager
from skit_bot import bot

# Создание класса, который будет представлять собой windows сервис
class AppServerSvc(win32serviceutil.ServiceFramework):
    # Имя сервиса, которое будет использоваться для управления сервисом
    _svc_name_ = "SKIT_bot"
    # Имя сервиса, которое будет отображаться в диспетчере служб сервиса Windows
    _svc_display_name_ = "SKIT_application_report"
    # Создание конструктора
    def __init__(self, args):
        # Вызов конструктора родительского класса
        win32serviceutil.ServiceFramework.__init__(self, args)
        # Создание события, которое будет использоваться для остановки сервиса
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        # Установка таймаута по умолчанию для сокетов
        socket.setdefaulttimeout(60)

    # Метод, который вызывается при остановке сервиса
    def SvcStop(self):
        # Сообщение о том, что сервис останавливается
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # Установка события для остановки сервиса
        win32event.SetEvent(self.hWaitStop)

    # Метод, который вызывается при запуске сервиса
    def SvcDoRun(self):
        # Логирование сообщений о запуске сервиса
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    # Основной метод сервиса, который выполнять основную логику
    def main(self):
        # Запуск бота в бесконечном цикле опроса
        while True:
            # Проверка на событие остановки сервиса
            rc = win32event.WaitForSingleObject(self.hWaitStop, 1000)
            if rc == win32event.WAIT_OBJECT_0:
                break
            bot.infinity_polling()


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(AppServerSvc)


