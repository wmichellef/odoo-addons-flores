# -*- coding: utf-8 -*-
from odoo import models, fields, api
from odoo.tools import format_date
from odoo.exceptions import UserError
import logging
from datetime import datetime, time
import io
import base64
import xlsxwriter

_logger = logging.getLogger(__name__)


class KardexCosteadoWizard(models.TransientModel):
    _name = 'kardex.costeado.wizard'
    _description = 'Kardex Costeado Report Wizard'

    date_from = fields.Date(string='Fecha Inicio', required=True)
    date_to = fields.Date(string='Fecha Fin', required=True)
    product_category_id = fields.Many2one('product.category', string='Categoría de Producto')
    product_ids = fields.Many2many('product.product', string='Productos', domain=[('type', '=', 'product')])
    location_id = fields.Many2one('stock.location', string='Ubicación')

    @api.onchange('product_category_id')
    def _onchange_product_category_id(self):
        if self.product_category_id:
            self.product_ids = self.env['product.product'].search([
                ('categ_id', '=', self.product_category_id.id),
                ('type', '=', 'product')
            ])
        else:
            self.product_ids = [(5, 0, 0)]

    @api.onchange('location_id')
    def _onchange_location_id(self):
        if self.location_id:
            quants = self.env['stock.quant'].search([
                ('location_id', '=', self.location_id.id),
                ('quantity', '>', 0),
                ('product_id.type', '=', 'product')
            ])
            product_ids = quants.mapped('product_id').ids
            self.product_ids = [(6, 0, product_ids)] if product_ids else [(5, 0, 0)]
        else:
            self.product_ids = [(5, 0, 0)]

    def action_generate_report(self):
        self.ensure_one()
        date_from_dt = datetime.combine(self.date_from, time(0, 0, 0))
        date_to_dt = datetime.combine(self.date_to, time(23, 59, 59))

        domain = [
            ('date', '>=', date_from_dt),
            ('date', '<=', date_to_dt),
            ('state', '=', 'done'),
        ]
        if self.product_ids:
            domain.append(('product_id', 'in', self.product_ids.ids))

        moves = self.env['stock.move'].search(domain, order='date, id')
        if not moves:
            data = {'kardex_data': {}, 'date_from': self.date_from, 'date_to': self.date_to}
        else:
            data = self._prepare_kardex_data(moves)
            data.update({'date_from': self.date_from, 'date_to': self.date_to})

        return {
            'type': 'ir.actions.report',
            'report_name': 'kardex_costeado_report.kardex_costeado_report_template',
            'report_type': 'qweb-pdf',
            'data': {
                'docs': self._ids,
                'kardex_data': data['kardex_data'],
                'date_from': data['date_from'].strftime('%Y-%m-%d'),
                'date_to': data['date_to'].strftime('%Y-%m-%d'),
            },
            'context': {'discard_logo_check': True},
        }

    def generate_excel_report(self):
        self.ensure_one()

        date_from_dt = datetime.combine(self.date_from, time(0, 0, 0))
        date_to_dt = datetime.combine(self.date_to, time(23, 59, 59))

        domain = [
            ('date', '>=', date_from_dt),
            ('date', '<=', date_to_dt),
            ('state', '=', 'done'),
        ]
        if self.product_ids:
            domain.append(('product_id', 'in', self.product_ids.ids))

        moves = self.env['stock.move'].search(domain, order='date, id')
        if not moves:
            raise UserError("No se encontraron movimientos en el rango de fechas seleccionado.")

        data = self._prepare_kardex_data(moves)
        kardex_data = data['kardex_data'] or {}

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        fmt_title = workbook.add_format({'bold': True, 'font_size': 14})
        fmt_subtitle = workbook.add_format({'bold': True, 'font_size': 12})
        fmt_th = workbook.add_format({'bold': True, 'bottom': 2, 'align': 'center', 'valign': 'vcenter'})
        fmt_th_left = workbook.add_format({'bold': True, 'bottom': 2, 'align': 'left'})
        fmt_right = workbook.add_format({'align': 'right'})
        fmt_right_bold = workbook.add_format({'align': 'right', 'bold': True})
        fmt_merge_th = workbook.add_format({'bold': True, 'bottom': 2, 'align': 'center', 'valign': 'vcenter'})
        fmt_product = workbook.add_format({'bold': True, 'font_size': 12})
        fmt_qty4 = workbook.add_format({'num_format': '#,##0.0000', 'align': 'right'})
        fmt_val2 = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'})
        fmt_unit4 = workbook.add_format({'num_format': '#,##0.0000', 'align': 'right'})
        fmt_blank = workbook.add_format({'align': 'left', 'valign': 'vcenter'})

        sheet = workbook.add_worksheet('Kardex Costeado')
        widths = [12, 12, 22, 18, 12, 14, 12, 14, 12, 14, 12, 12, 14, 12]
        for i, w in enumerate(widths):
            sheet.set_column(i, i, w)

        row = 0
        sheet.write(row, 0, "Kardex Costeado Report", fmt_title); row += 1
        sheet.write(row, 0, f"Período: {self.date_from} - {self.date_to}", fmt_subtitle); row += 2

        for product_dict in kardex_data.values():
            product_name = product_dict.get('product') or product_dict.get('product_name') or ''
            sheet.write(row, 0, f"Producto: {product_name}", fmt_product)
            row += 1

            sheet.write(row, 0, "Fecha", fmt_th_left)
            sheet.write(row, 1, "Tipo", fmt_th_left)
            sheet.write(row, 2, "Documento", fmt_th_left)
            sheet.write(row, 3, "Bodega", fmt_th_left)
            sheet.merge_range(row, 4, row, 5, "Existencia Anterior", fmt_merge_th)
            sheet.merge_range(row, 6, row, 7, "Entradas", fmt_merge_th)
            sheet.merge_range(row, 8, row, 9, "Salidas", fmt_merge_th)
            sheet.write(row, 10, "Traslado", fmt_th)
            sheet.merge_range(row, 11, row, 12, "Existencia Actual", fmt_merge_th)
            sheet.write(row, 13, "Costo Actual", fmt_right_bold)
            row += 1

            for c in range(0, 4):
                sheet.write_blank(row, c, None, fmt_blank)
            sheet.write(row, 4, "Unidades", fmt_right_bold)
            sheet.write(row, 5, "Valor", fmt_right_bold)
            sheet.write(row, 6, "Unidades", fmt_right_bold)
            sheet.write(row, 7, "Valor", fmt_right_bold)
            sheet.write(row, 8, "Unidades", fmt_right_bold)
            sheet.write(row, 9, "Valor", fmt_right_bold)
            sheet.write(row, 10, "Unidades", fmt_right_bold)
            sheet.write(row, 11, "Unidades", fmt_right_bold)
            sheet.write(row, 12, "Valor", fmt_right_bold)
            sheet.write(row, 13, "(Unit)", fmt_right_bold)
            row += 1

            total_in_qty = total_in_value = total_out_qty = total_out_value = total_transfer_qty = 0.0
            first_row = True
            disp_prev_qty = disp_prev_value = 0.0
            final_qty_eff = final_value_eff = 0.0

            for mv in product_dict.get('moves', []):
                ref = mv.get('ref') or ''
                t = mv.get('type') or ''
                is_in = (t == 'Entrada')
                is_out = (t == 'Salida')
                is_internal = ('WH/INT/' in ref) or (t in ['Traslado', 'Transfer'])

                prev_qty_line = (mv.get('previous_qty') or 0.0) if first_row else disp_prev_qty
                prev_val_line = (mv.get('previous_value') or 0.0) if first_row else disp_prev_value

                qty = mv.get('qty') or 0.0
                total_move = mv.get('total_move') or 0.0
                in_qty_line = qty if (not is_internal and is_in) else 0.0
                in_val_line = total_move if (not is_internal and is_in) else 0.0
                if mv.get('is_cost_adjustment'):
                    in_val_line = total_move
                    in_qty_line = 0.0
                out_qty_line = qty if (not is_internal and is_out) else 0.0
                out_val_line = total_move if (not is_internal and is_out) else 0.0
                transfer_qty_line = qty if is_internal else 0.0

                balance_qty_disp = prev_qty_line + in_qty_line - out_qty_line
                balance_value_disp = prev_val_line + in_val_line - out_val_line
                unit_cost_line = (balance_value_disp / balance_qty_disp) if balance_qty_disp else 0.0

                total_in_qty += in_qty_line
                total_in_value += in_val_line
                total_out_qty += out_qty_line
                total_out_value += out_val_line
                total_transfer_qty += transfer_qty_line

                if not is_internal:
                    final_qty_eff = balance_qty_disp
                    final_value_eff = balance_value_disp

                disp_prev_qty = balance_qty_disp
                disp_prev_value = balance_value_disp
                first_row = False

                sheet.write(row, 0, mv.get('date') or '')
                sheet.write(row, 1, t)
                sheet.write(row, 2, ref)
                sheet.write(row, 3, mv.get('warehouse') or '')
                sheet.write_number(row, 4, prev_qty_line, fmt_qty4)
                sheet.write_number(row, 5, prev_val_line, fmt_val2)
                sheet.write_number(row, 6, in_qty_line, fmt_qty4)
                sheet.write_number(row, 7, in_val_line, fmt_val2)
                sheet.write_number(row, 8, out_qty_line, fmt_qty4)
                sheet.write_number(row, 9, out_val_line, fmt_val2)
                sheet.write_number(row, 10, transfer_qty_line, fmt_qty4)
                sheet.write_number(row, 11, balance_qty_disp, fmt_qty4)
                sheet.write_number(row, 12, balance_value_disp, fmt_val2)
                sheet.write_number(row, 13, unit_cost_line, fmt_unit4)
                row += 1

            sheet.write(row, 0, "Total Entradas:", fmt_right_bold)
            sheet.write_number(row, 1, total_in_qty, fmt_qty4)
            sheet.write(row, 2, "Valor:", fmt_right)
            sheet.write_number(row, 3, total_in_value, fmt_val2)
            row += 1

            sheet.write(row, 0, "Total Salidas:", fmt_right_bold)
            sheet.write_number(row, 1, total_out_qty, fmt_qty4)
            sheet.write(row, 2, "Valor:", fmt_right)
            sheet.write_number(row, 3, total_out_value, fmt_val2)
            row += 1

            sheet.write(row, 0, "Total Traslados:", fmt_right_bold)
            sheet.write_number(row, 1, total_transfer_qty, fmt_qty4)
            row += 1

            sheet.write(row, 0, "Saldo Final:", fmt_right_bold)
            sheet.write_number(row, 1, final_qty_eff or 0.0, fmt_qty4)
            sheet.write(row, 2, "Valor:", fmt_right)
            sheet.write_number(row, 3, final_value_eff or 0.0, fmt_val2)
            row += 2

        workbook.close()
        output.seek(0)

        attachment = self.env['ir.attachment'].create({
            'name': 'Kardex_Costeado.xlsx',
            'type': 'binary',
            'datas': base64.b64encode(output.read()),
            'res_model': self._name,
            'res_id': self.id,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })

        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/%s?download=true' % attachment.id,
            'target': 'self',
        }

    
    def _prepare_kardex_data(self, moves):
            """
            Construye el kardex considerando:
            - Movimientos de stock (stock.move) del rango
            - Ajustes de costo (stock.valuation.layer con quantity=0 y value!=0 y SIN stock_move)
            """
            kardex_data = {}
            products = moves.mapped('product_id')
            date_from_dt = datetime.combine(self.date_from, time(0, 0, 0))
            date_to_dt = datetime.combine(self.date_to, time(23, 59, 59))

            for product in products:
                product_moves = moves.filtered(lambda m: m.product_id == product).sorted(key=lambda m: (m.date, m.id))
                kardex_data[product.id] = {
                    'product': f"[{product.default_code or ''}] {product.name or ''}",
                    'moves': [],
                    'initial_qty': 0.0,
                    'initial_value': 0.0,
                }

                # Saldos de apertura hasta la 00:00 del date_from
                self.env.cr.execute("""
                    SELECT
                        COALESCE(SUM(svl.quantity), 0.0) AS qty,
                        COALESCE(SUM(svl.value), 0.0)    AS val
                    FROM stock_valuation_layer svl
                    WHERE svl.product_id = %s
                      AND svl.create_date < %s
                """, (product.id, date_from_dt))
                row = self.env.cr.fetchone() or (0.0, 0.0)
                prev_qty = row[0] or 0.0
                prev_value = row[1] or 0.0
                kardex_data[product.id]['initial_qty'] = prev_qty
                kardex_data[product.id]['initial_value'] = prev_value

                # SVL vinculados a movimientos del rango
                layers = self.env['stock.valuation.layer'].search([
                    ('stock_move_id', 'in', product_moves.ids),
                    ('product_id', '=', product.id),
                ])
                svl_map = {}
                for l in layers:
                    d = svl_map.setdefault(l.stock_move_id.id, {'qty': 0.0, 'val': 0.0})
                    d['qty'] += (l.quantity or 0.0)
                    d['val'] += (l.value or 0.0)

                def _unit_cost_from(qty, val):
                    return (val / qty) if qty else 0.0

                # Ajustes de costo del rango (sin stock_move_id)
                adj_layers = self.env['stock.valuation.layer'].search([
                    ('stock_move_id', '=', False),
                    ('product_id', '=', product.id),
                    ('create_date', '>=', date_from_dt),
                    ('create_date', '<=', date_to_dt),
                    ('value', '!=', 0.0),
                ])

                # Unir eventos: movimientos + ajustes
                events = [{'type': 'move', 'date': m.date, 'id': m.id, 'move': m} for m in product_moves]
                events += [{'type': 'adjust', 'date': a.create_date, 'id': a.id, 'adj': a} for a in adj_layers]
                events.sort(key=lambda e: (e['date'], e['id']))

                for ev in events:
                    if ev['type'] == 'move':
                        m = ev['move']
                        src = m.location_id.usage
                        dst = m.location_dest_id.usage
                        if src == 'internal' and dst == 'internal':
                            t = 'Traslado'
                        elif dst == 'internal':
                            t = 'Entrada'
                        elif src == 'internal':
                            t = 'Salida'
                        else:
                            t = 'Otro'

                        svl = svl_map.get(m.id)
                        mv_qty = svl['qty'] if svl else 0.0
                        mv_val = svl['val'] if svl else 0.0
                        is_in = (dst == 'internal') and (src != 'internal')
                        is_out = (src == 'internal') and (dst != 'internal')

                        qty_abs = abs(mv_qty) if mv_qty else abs(m.product_uom_qty if is_out else (m.quantity_done or 0.0))
                        total_abs = abs(mv_val)

                        # Calcular costo unitario preferentemente desde SVL
                        unit_cost = _unit_cost_from(abs(mv_qty), mv_val) if mv_qty else _unit_cost_from(prev_qty if is_out else qty_abs, mv_val if mv_val else prev_value if is_out else 0.0)

                        # Actualizar saldos
                        if is_in:
                            bal_qty = prev_qty + qty_abs
                            bal_val = prev_value + total_abs
                        elif is_out:
                            bal_qty = prev_qty - qty_abs
                            bal_val = prev_value - total_abs
                        else:  # traslado u otro interno a interno usa valores acumulados
                            bal_qty = prev_qty
                            bal_val = prev_value

                        kardex_data[product.id]['moves'].append({
                            'date': m.date,
                            'type': t,
                            'document': m.reference or m.name,
                            'warehouse': (m.location_dest_id.warehouse_id.name if is_in else m.location_id.warehouse_id.name) or '',
                            'previous_qty': prev_qty,
                            'previous_value': prev_value,
                            'qty': qty_abs,
                            'unit_cost': unit_cost,
                            'total_move': total_abs,
                            'balance_qty': bal_qty,
                            'balance_value': bal_val, 'ref': (m.reference or m.name or ''), 'warehouse': (m.location_id.complete_name or m.location_id.display_name or m.location_id.name or ''),
                        })
                        prev_qty = bal_qty
                        prev_value = bal_val

                    else:  # adjust
                        a = ev['adj']
                        # Ajuste no modifica cantidad, solo valor
                        qty_abs = 0.0
                        total_abs = abs(a.value or 0.0)
                        bal_qty = prev_qty
                        bal_val = prev_value + (a.value or 0.0)
                        unit_cost = _unit_cost_from(bal_qty, bal_val) if bal_qty else 0.0

                        kardex_data[product.id]['moves'].append({
                            'date': a.create_date,
                            'type': 'Ajuste de costo',
                            'document': a.description or 'Ajuste de costo',
                            'warehouse': '',
                            'previous_qty': prev_qty,
                            'previous_value': prev_value,
                            'qty': qty_abs,
                            'unit_cost': unit_cost,
                            'total_move': total_abs,
                            'balance_qty': bal_qty,
                            'balance_value': bal_val,
                            'is_cost_adjustment': True,
                            'adjust_value': a.value,
                        })
                        prev_qty = bal_qty
                        prev_value = bal_val

            return {'kardex_data': kardex_data}