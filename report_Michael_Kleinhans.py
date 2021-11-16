from openerp.report.report_sxw import report_sxw
from openerp import pooler

import cStringIO
import xlwt
import logging
_logger = logging.getLogger(__name__)

class report_xls(report_sxw):

    def create(self, cr, uid, ids, data, context=None):
        self.pool = pooler.get_pool(cr.dbname)
        self.cr = cr
        self.uid = uid
        report_obj = self.pool.get('ir.actions.report.xml')
        report_ids = report_obj.search(
            cr, uid, [('name', '=', 'Kleinhans_xls')], context=context)
        if report_ids:
            report_xml = report_obj.browse(
                cr, uid, report_ids[0], context=context)
            self.title = report_xml.name
            if report_xml.report_type == 'xls':
                return self.create_source_xls(cr, uid, ids, data, context)
        return 0

    def create_source_xls(self, cr, uid, ids, data, context):
        purchase_order_obj = self.pool.get('purchase.order')
        purchase_order_ids = purchase_order_obj.search(
           cr, uid, [('id', '=', ids[0])], context=context)
        browsed_purchase_order = purchase_order_obj.browse(
            cr, uid, purchase_order_ids[0], context=context)
        csv_header = ["procurement_group_id","display_name","product_name", "Description for supplier", "Mold #", "Additional comments", "product_code","product_qty", "production_price", "picking_type_id","write_date", "shipping_country", "engraving", "Stein_1", "Stein_2"]
        pol_obj = self.pool.get('purchase.order.line')
        pol_ids = pol_obj.search(
            cr, uid, [('order_id', '=', ids[0])], context=context)

        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet("sheet")
        j = 0
        for i in csv_header:
            worksheet.write(0, j, i)
            j = j +1

        for po in purchase_order_ids:
            browsed_po = purchase_order_obj.browse(
                cr, uid, po, context=context)

        k = 1
        for line in pol_ids:
            line_content = dict.fromkeys(["display_name",
                                          "product_name",
                                          "Description for supplier",
                                          "Mold #",
                                          "Additional comments",
                                          "product_code",
                                          "product_qty",    
                                          "production_price",
                                          "picking_type_id",
                                          "write_date",
                                          "procurement_group_id",
                                          "shipping_country",
                                          "engraving",
					  "Stein_1",
					  "Stein_2"])
            browsed_line = pol_obj.browse(
                cr, uid, line, context=context)
            line_content["display_name"] = browsed_purchase_order.display_name
            line_content["product_name"] = "\"" + browsed_line.product_id.name
            if browsed_line.product_id.attribute_value_ids is not False:
                for attr in browsed_line.product_id.attribute_value_ids:
                    line_content["product_name"] += '- ' + attr.name
            line_content["product_name"] += "\""
            # line_content["product_name"] = browsed_line.product_id.name
            if browsed_line.product_id.description_purchase_variant is not False:
                line_content["Description for supplier"] = browsed_line.product_id.description_purchase_variant
            else:
                line_content["Description for supplier"] = None
            line_content["Mold #"] = browsed_line.product_id.seller_ids[0].product_code
            line_content["product_code"] = browsed_line.product_id.code
            line_content["product_qty"] = browsed_line.product_qty
            line_content["production_price"] = browsed_line.product_id.standard_price
            line_content["picking_type_id"] = browsed_purchase_order.picking_type_id.complete_name
            line_content["write_date"] = browsed_line.write_date
            line_content["shipping_country"] = " "

            position_marker = browsed_line.name.find('comment:')
            if position_marker != -1:
                comments = browsed_line.name[position_marker + 8:]
                line_content["Additional comments"] = comments
            else:
                line_content["Additional comments"] = None

            if 'Engraving' in browsed_line.name:
                mark_1 = browsed_line.name.find("Engraving: ")
                rest_name = browsed_line.name[mark_1 + 11:]
                mark_2 = rest_name.find(" ||")
                line_content["engraving"] = rest_name[:mark_2]
            elif 'Three pendants' in browsed_line.name:
                mark_1 = browsed_line.name.find("Three pendants: ")
                rest_name = browsed_line.name[mark_1 + 16:]
                mark_2 = rest_name.find(" ||")
                line_content["engraving"] = rest_name[:mark_2]
            else:
                line_content["engraving"] = " "

            if 'Stein_1' in browsed_line.name:
                mark_1 = browsed_line.name.find("Stein_1: ")
                rest_name = browsed_line.name[mark_1 + 9:]
                mark_2 = rest_name.find(" ||")
                line_content["Stein_1"] = rest_name[:mark_2]
            else:
                line_content["Stein_1"] = " "

            if 'Stein_2' in browsed_line.name:
                mark_1 = browsed_line.name.find("Stein_2: ")
                rest_name = browsed_line.name[mark_1 + 9:]
                mark_2 = rest_name.find(" ||")
                line_content["Stein_2"] = rest_name[:mark_2]
            else:
                line_content["Stein_2"] = " "
            

            total_qty = browsed_line.product_qty
            if browsed_line.procurement_ids:
                so_qty = 0
                line_content["product_qty"] = 1
                
                for group in browsed_line.procurement_ids:
                    order_name = group.origin.split(':')[0]
                    try:
                        id_order = self.pool.get('sale.order').search(cr, uid, [('name', '=', order_name)], context=context)[0]
                    except IndexError:
                        continue
                    if self.pool.get('sale.order').browse(cr, uid, id_order, context=context)[0].state != 'cancel':
                        for proc in group.group_id.procurement_ids:
                            if proc.sale_line_id.order_id.partner_shipping_id.country_id.name != False:
                                line_content["shipping_country"] = proc.sale_line_id.order_id.partner_shipping_id.country_id.name



                        so_qty = so_qty + group.product_qty
                        for i in range(0, int(group.product_qty)):
                            line_content["procurement_group_id"] = group.group_id.display_name
                            self.xls_line(line_content, worksheet, k)
                            k = k + 1
                remaining_qty = total_qty - so_qty
                if remaining_qty > 0:
                    line_content["procurement_group_id"] = " "
                    line_content["shipping_country"] = " "
                    for i in range(0, int(remaining_qty)):
                        self.xls_line(line_content, worksheet, k)
                        k = k + 1

            else:
                line_content["product_qty"] = 1
                line_content["procurement_group_id"] = " "
                for i in range(0, int(browsed_line.product_qty)):
                    self.xls_line(line_content, worksheet, k)
                    k = k + 1


            n = cStringIO.StringIO()
            workbook.save(n)
            n.seek(0)
        return n.read(), "xls"

    def xls_line(self, line_content, worksheet, k):
        header_order = ["procurement_group_id",
                        "display_name",
                        "product_name",
                        "Description for supplier",
                        "Mold #",
                        "Additional comments",
                        "product_code",
                        "product_qty",
                        "production_price",
                        "picking_type_id",
                        "write_date",
                        "shipping_country",
                        "engraving",
			"Stein_1",
			"Stein_2"]
        i = 0
        for f in header_order:
            worksheet.write(k, i, line_content[f])
            i = i + 1
