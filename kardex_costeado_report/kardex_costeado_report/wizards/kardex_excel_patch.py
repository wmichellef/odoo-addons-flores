
# -*- coding: utf-8 -*-
from odoo import models
from odoo.exceptions import UserError
from datetime import datetime, time
import io
import base64
import xlsxwriter

class KardexCosteadoWizardAdjust(models.TransientModel):
    _inherit = 'kardex.costeado.wizard'

    def generate_excel_report(self):
        # Re-implement with single 'Ajuste' (value) column next to 'Traslado'
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
        kardex_data = data.get('kardex_data') or {}

        # Si el _prepare_kardex_data base no trae flags de ajuste, no rompemos: seteamos default
        def _mv_is_adj(mv): return bool(mv.get('is_adjustment'))
        def _mv_adj_val(mv): 
            v = mv.get('adjustment_value') or 0.0
            try:
                return float(v)
            except Exception:
                return 0.0

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
        # Column layout: Fecha, Tipo, Documento, Bodega, Existencia Anterior (2),
        # Entradas (2), Salidas (2), Traslado (Unidades), Ajuste (Valor), Existencia Actual (2), Costo Actual
        widths = [12, 12, 22, 18, 12, 14, 12, 14, 12, 14, 12, 14, 12, 12, 14]
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
            sheet.write(row, 11, "Ajuste", fmt_right_bold)  # NUEVO: solo valor
            sheet.merge_range(row, 12, row, 13, "Existencia Actual", fmt_merge_th)
            sheet.write(row, 14, "Costo Actual", fmt_right_bold)
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
            sheet.write(row, 11, "Valor", fmt_right_bold)      # Ajuste (valor)
            sheet.write(row, 12, "Unidades", fmt_right_bold)
            sheet.write(row, 13, "Valor", fmt_right_bold)
            sheet.write(row, 14, "(Unit)", fmt_right_bold)
            row += 1

            first_row = True
            disp_prev_qty = disp_prev_value = 0.0

            for mv in product_dict.get('moves', []):
                # Saldos previos
                prev_qty_line = (mv.get('previous_qty') or 0.0) if first_row else disp_prev_qty
                prev_val_line = (mv.get('previous_value') or 0.0) if first_row else disp_prev_value

                t = mv.get('type') or ''
                qty = mv.get('qty') or 0.0
                total_move = mv.get('total_move') or 0.0
                is_internal = (t == 'Traslado')
                is_in = (t == 'Entrada')
                is_out = (t == 'Salida')
                is_adjustment = _mv_is_adj(mv)
                adj_val_line = _mv_adj_val(mv) if is_adjustment else 0.0

                in_qty_line = qty if (not is_internal and is_in and not is_adjustment) else 0.0
                in_val_line = total_move if (not is_internal and is_in and not is_adjustment) else 0.0
                out_qty_line = qty if (not is_internal and is_out and not is_adjustment) else 0.0
                out_val_line = total_move if (not is_internal and is_out and not is_adjustment) else 0.0
                transfer_qty_line = qty if (is_internal and not is_adjustment) else 0.0

                # Saldos mostrados
                if is_adjustment:
                    balance_qty_disp = prev_qty_line
                    balance_value_disp = prev_val_line + adj_val_line  # solo valor
                else:
                    balance_qty_disp = prev_qty_line + in_qty_line - out_qty_line
                    balance_value_disp = prev_val_line + in_val_line - out_val_line

                unit_cost_line = (balance_value_disp / balance_qty_disp) if balance_qty_disp else 0.0

                # Pintar fila
                from odoo.tools import format_date as _fmt_date
                date_text = mv.get('date') or ''
                sheet.write(row, 0, date_text)
                # Forzar texto "Ajuste" cuando es ajuste
                sheet.write(row, 1, ("Ajuste" if is_adjustment else t))
                sheet.write(row, 2, mv.get('ref') or '')
                sheet.write(row, 3, mv.get('warehouse') or '')
                sheet.write_number(row, 4, prev_qty_line, fmt_qty4)
                sheet.write_number(row, 5, prev_val_line, fmt_val2)
                sheet.write_number(row, 6, in_qty_line, fmt_qty4)
                sheet.write_number(row, 7, in_val_line, fmt_val2)
                sheet.write_number(row, 8, out_qty_line, fmt_qty4)
                sheet.write_number(row, 9, out_val_line, fmt_val2)
                sheet.write_number(row, 10, transfer_qty_line, fmt_qty4)
                sheet.write_number(row, 11, (adj_val_line if is_adjustment else 0.0), fmt_val2)  # Ajuste (sin signo)
                sheet.write_number(row, 12, balance_qty_disp, fmt_qty4)
                sheet.write_number(row, 13, balance_value_disp, fmt_val2)
                sheet.write_number(row, 14, unit_cost_line, fmt_unit4)
                row += 1

                first_row = False
                disp_prev_qty = balance_qty_disp
                disp_prev_value = balance_value_disp

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
