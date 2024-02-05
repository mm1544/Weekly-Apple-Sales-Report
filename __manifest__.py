{
    'name': 'Weekly Apple Sales Report',
    'version': '1.0',
    'category': 'Generic Modules/Others',
    'summary': 'Generates Weekly Apple Sales Report XLSX and sends to the designated person.',
    'sequence': '1',
    'author': 'Martynas Minskis',
    'depends': ['sale'],
    'demo': [],
    'data': [

        #        Sequence: security, data, wizards, views
        'views/weekly_apple_sales_report.xml',
    ],
    'demo': [],
    'qweb': [],

    'installable': True,
    'application': True,
    'auto_install': False,
    #     'licence': 'OPL-1',
}
