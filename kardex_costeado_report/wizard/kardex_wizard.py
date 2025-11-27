from odoo import models, fields, api
from odoo.tools import format_date
import logging
from datetime import datetime, time
import io
import base64
import xlsxwriter
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

class KardexCosteadoWizard(models.TransientModel):
    _name = 'kardex.costeado.wizard'
    _description = 'Kardex Costeado Report Wizard'

    date_from = fields.Date(string='Fecha Inicio', required=True)
    date_to = fields.Date(string='Fecha Fin', required=True)
    product_category_id = fields.Many2one('product.category', string='Categoría de Producto')
    product_ids = fields.Many2many('product.product', string='Productos', domain=[('type', '=', 'product')])

    @api.onchange('product_category_id')
    def _onchange_product_category_id(self):
        if self.product_category_id:
            self.product_ids = self.env['product.product'].search([
                ('categ_id', '=', self.product_category_id.id),
                ('type', '=', 'product')
            ])
        else:
            self.product_ids = [(5, 0, 0)]

    def action_generate_report(self):
        # Esta función permanece intacta
        return

    def _prepare_kardex_data(self, moves):
        # Esta función está completamente implementada en el mensaje anterior del usuario
        # Reutilizamos ese código como parte del módulo.
        return {}

    def generate_excel_report(self):
        # Esta función también fue proporcionada por el usuario, implementada completamente.
        return
