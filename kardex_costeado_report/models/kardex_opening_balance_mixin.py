# -*- coding: utf-8 -*-
from odoo import models, api
from odoo.tools import float_round

class KardexOpeningBalanceMixin(models.AbstractModel):
    _name = "kardex.opening.balance.mixin"
    _description = "Utilidades para calcular existencia anterior (qty y valor)"

    @api.model
    def _kardex_get_opening_balance(self, product_id, location_ids, company_id, date_from, uom=None, rounding=4):
        """ Devuelve (qty, value) acumulado de stock.valuation.layer hasta date_from - 1 microsegundo.
            - product_id: int
            - location_ids: lista de ids de ubicaciones (o None para todas de la compañía)
            - company_id: int
            - date_from: datetime (inicio del rango seleccionado por el usuario)
            - uom: record de product.uom para convertir qty (opcional)
        """
        if not date_from:
            return 0.0, 0.0

        cutoff = date_from
        cutoff_str = cutoff.strftime('%Y-%m-%d %H:%M:%S')

        # Nota: SVL no siempre guarda location_*; para generalidad usamos valor acumulado por layers del producto/compañía.
        self.env.cr.execute(
            "SELECT COALESCE(SUM(quantity), 0.0) AS qty, COALESCE(SUM(value), 0.0) AS val "
            "FROM stock_valuation_layer "
            "WHERE product_id = %s AND company_id = %s AND create_date < %s",
            (product_id, company_id, cutoff_str)
        )
        res = self.env.cr.fetchone() or (0.0, 0.0)
        qty = res[0] or 0.0
        val = res[1] or 0.0

        if uom:
            prod = self.env['product.product'].browse(product_id)
            qty = prod.uom_id._compute_quantity(qty, uom)

        qty = float_round(qty, precision_digits=rounding)
        val = float_round(val, precision_digits=2)
        return qty, val
