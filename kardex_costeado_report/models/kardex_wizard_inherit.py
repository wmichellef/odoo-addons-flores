# -*- coding: utf-8 -*-
from odoo import api, models, fields

class KardexCosteadoWizardInherit(models.TransientModel):
    _inherit = "kardex.costeado.wizard"  # Ajusta si tu wizard tiene otro _name

    opening_qty = fields.Float(string="Existencia Anterior (Unid)", readonly=True)
    opening_val = fields.Float(string="Existencia Anterior (Valor)", readonly=True)

    def _compute_opening_for_product(self, product, date_from, locations=None):
        mixin = self.env['kardex.opening.balance.mixin']
        location_ids = locations and locations.ids or None
        qty, val = mixin._kardex_get_opening_balance(
            product_id=product.id,
            location_ids=location_ids,
            company_id=self.env.company.id,
            date_from=fields.Datetime.to_datetime(date_from),
            uom=product.uom_id,
            rounding=4,
        )
        return qty, val

    def _prepare_kardex_data(self, moves):
        data = super()._prepare_kardex_data(moves)

        kardex_data = {}
        try:
            kardex_data = data.get('kardex_data') or {}
        except Exception:
            pass

        date_from = getattr(self, 'date_from', None)
        if not date_from:
            return data

        locations = getattr(self, 'location_ids', False)

        for prod_key, prod_dict in kardex_data.items():
            product = None
            if isinstance(prod_key, int):
                product = self.env['product.product'].browse(prod_key)
            else:
                product = self.env['product.product'].search([('name', '=', prod_key)], limit=1) or \
                          self.env['product.product'].search([('default_code', '=', prod_key)], limit=1)
            if not product and isinstance(prod_dict, dict):
                pid = prod_dict.get('product_id')
                if pid:
                    product = self.env['product.product'].browse(pid)

            if not product:
                continue

            qty, val = self._compute_opening_for_product(product, date_from, locations=locations)
            if isinstance(prod_dict, dict):
                prod_dict['opening_qty'] = qty
                prod_dict['opening_val'] = val

        data['kardex_data'] = kardex_data
        return data
