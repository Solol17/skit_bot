Проект "Skit_bot".
Авторы проекта: Евгений Сокол и Олег Солдатов.
Skit_bot - tg-бот, предназначенный для оперативных сообщений о количестве заявок групп ЦОП в системе СКИТ, а также предоставление данной информации в виде отчета файла .docx.
Для установки на машину, требуется:
1. Развернуть 2 docker-контейнера: skit_bot и redis.
2. Добавить в директорию, где расположен skit_bot, файл .env, в котором указано:
   a) Логин от СКИТ;
   б) Пароль от СКИТ;
   в) Токен tg-бота;
   г) Ключ для использования Mistral.
3. Добавить в директорию, где расположен skit_bot, geckodriver*.
* - если будут требоваться варианты разворачивания проекта на windows, без использования docker, следует использовать соответствующую версию geckodriver для windows.
4. Телегам-бот с названием "Заявки ЦОП" принимает 2 сообщения: /start и /report.
   /start - предоставляет информацию в виде сообщений;
   /report - предоставляет информацию в виде .docx файла.