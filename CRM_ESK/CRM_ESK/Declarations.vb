Module Declarations
    Public CompanyID As String                            'строка - номер компании в Scala
    Public Year As String                                 'cтрока - год в Scala
    Public UserCode As String                             'короткое имя пользователя Scala
    Public UserID As Integer                              'код пользователя Scala
    Public OwnerID As Integer                             'код владельца записи Scala
    Public FullName As String                             'ФИО пользователя Scala
    Public CC As String                                   'кост центр пользователя Scala
    Public SalesmanCode As String                         'код продавца в Скала

    Public MyConnStr As String                            'строка соединения с БД
    Public MyNETConnStr As String                         '.NET строка соединения с БД
    Public MyConn As ADODB.Connection                     'соединение с БД
    Public MyRec As ADODB.Recordset                       'Рекордсет он в Африке рекордсет

    Public MyCCPermission As Boolean                      'Входит  или нет в группу CRMManagers
    Public MyPermission As Boolean                        'Входит  или нет в группу CRMDirector
    Public MyPDPermission As Boolean                      'Входит  или нет в группу ProjectDirector
    Public AllowChangeUser As String                      'Разрешено ли пользователям из групп CRMManagers и CRMDirector менять 

    Public MyResult As Integer                            'результат выполнения операции
    Public MyEventID As String                            'ID действия
    Public MyOldEventID As String                         'ID действия старый - для копирования
    Public MyClientID As String                           'ID клиента
    Public MyNewClientID As String                        'Новый ID клиента (при объединении)
    Public MyContactID As String                          'ID контакта
    Public MyProjectID As String                          'ID (GUID) проекта
    Public MyParentProjectID As String                    'ID (GUID) родительского проекта
    Public MyAttachmentID As String                       'IDаттачмента
    Public MyOrderID As String                            'ID (GUID) заказа на продажу
    Public MyHighPriceID As String                        'ID записи в High Price
    Public MyWHAbsencesID As String                       'ID записи в WHAbsences

    Public MyAddEvent As AddEvent                         'Реализация окна добавления / редактирования действия
    Public MyCustomerSelect As CustomerSelect             'реализация окна поиска клиентов
    Public MyContactSelect As ContactSelect               'Реализациия окна выбора контактов
    Public MySalesOrderList As SalesOrderList             'реализация окна ввода заказов на продажу
    Public MyHighPrice As HighPrice                       'Реализация окна добавления запасов с высокой ценой
    Public MyWHAbsences As WHAbsences                     'реализация окна добавления информации по товарам с недостаточным количеством
    Public MyProjectSelect As ProjectSelect               'реализация окна выбора проекта для клиента
    Public MyCustomerSelectList As CustomerSelectList     'реализация окна выборки при поиске клиентов
    Public MyAddClient As AddClient                       'Реализация окна добавления / редактирования клиента
    Public MyCustomerMerge As CustomerMerge               'реализация окна объединения клиентов
    Public MyCustomerExtInfo As CustomerExtInfo           'реализация окна ввода дополнительной информации по клиенту
    Public MyAddContact As AddContact                     'Реализация окна добавления / редактирования контакта
    Public MyAddProject As AddProject                     'реализация окна создания / редактирования проекта
    Public MyViewEvent As ViewEvent                       'Реализация окна просмотра действия
    Public MyProjectList As ProjectList                   'реализация окна выбора проекта 
    Public MyProjectDetailsImport As ProjectDetailsImport 'реализация окна импорта детальной инфоормации по проекту
    Public MyItemGroupsInProject As ItemGroupsInProject   'реализация окна со списком групп товаров
    Public MyPlansApprovement As PlansApprovement         'Реализация окна утверждения планов
    Public MySendReturnAction As SendReturnAction         'Реализация окна передачи /возврата действий CRM
    Public MyAddSalesOrder As AddSalesOrder               'реализация окна добавления заказа на продажу
    Public MySupplierSelect As SupplierSelect             'реализация окна поиска поставщиков
    Public MySupplierSelectList As SupplierSelectList     'реализация окна выборки при поиске поставщиков
    Public MyItemSelectList As ItemSelectList             'реализация окна выборки при поиске запасов
    Public MyItemHighPrice As ItemHighPrice               'реализация окна добавления информации по товару с высокой ценой
    Public MyItemWHAbsences As ItemWHAbsences             'реализация окна добавления информации по товару с недостаточным количеством
End Module
