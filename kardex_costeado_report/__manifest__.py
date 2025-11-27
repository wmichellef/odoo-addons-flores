{
    'name': 'Kardex Costeado Report',
    'version': '1.0',
    'summary': 'Custom Kardex report with cost details for Odoo 17',
    'author': 'Wens',
    'category': 'Inventory',
    'depends': ['stock', 'stock_account'],
    'data': [
        'security/ir.model.access.csv',
        'wizards/kardex_wizard_views.xml',
        'reports/kardex_report.xml',
        'reports/kardex_report_template.xml',
    ],
    'installable': True,
    'application': False,
}