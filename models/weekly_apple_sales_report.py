from odoo import models, fields, api
from datetime import datetime, timedelta, date
import xlsxwriter
import io
import base64
import logging
import re

_logger = logging.getLogger(__name__)


class WeeklyAppleSalesReport(models.Model):
    _inherit = 'account.move'

    xlsx_report_file = fields.Binary(
        string='XLSX Report File', attachment=True)

    def get_previous_week_monday_date(self, today):
        """Return the date of previous week's Monday"""

        # For Monday today.weekday()=0. For Sunday today.weekday()=6.
        days_to_subtract = today.weekday() % 7 or 7
        if today.weekday() > 0:  # If today is not Monday
            days_to_subtract += 7

        return today - timedelta(days=days_to_subtract)

    def get_address(self, contact):
        """Return formatted address from a contact"""
        address_parts = [contact.street, contact.street2, contact.x_address3]
        return ', '.join(filter(None, address_parts))

    def get_and_process_invoice_data(self):
        """Finds Invoices and returns data"""
        # OK / TO CHANGE IN PRODUCTION
        today_date = datetime.today()
        # TEST
        # today_date = datetime(2023,11,30)
        # today_date = datetime(2023,12,7)

        previous_week_monday_date = self.get_previous_week_monday_date(
            today_date)

        # start_of_previous_week: 2023-12-25 00:00:00
        start_of_previous_week = datetime(
            previous_week_monday_date.year, previous_week_monday_date.month, previous_week_monday_date.day)

        # start_of_this_week: 2024-01-01 00:00:00
        start_of_this_week = start_of_previous_week + timedelta(days=7)

        invoices = self.env['account.move'].search([
            ('invoice_date', '>=', start_of_previous_week),
            ('invoice_date', '<', start_of_this_week),
            ('state', 'in', ['posted']),
            ('move_type', '=', 'out_invoice'),
            #  The domain [('invoice_line_ids.product_id.test_field', '=', True)] checks each invoice line associated with an invoice.
            # It filters the invoices where at least one of the invoice lines has a product with x_include_in_apple_s2w_report set to True.
            ('invoice_line_ids.product_id.x_include_in_apple_s2w_report', '=', True),
        ])

        if not invoices:
            # raise Warning('No invoices found')
            _logger.warning('WARNING: No invoice found')
            return []

        result = [['Product Code', 'Invoice Quantity', 'Invoice Number', 'Invoice Date',
                   'Delivery Address', 'Address', 'Town/City', 'Country', 'School or Business']]

        for invoice in invoices:
            for invoice_line in invoice.invoice_line_ids.filtered(lambda l: l.product_id.x_include_in_apple_s2w_report):
                result.append([
                    invoice_line.product_id.default_code,
                    invoice_line.quantity,
                    invoice.name,
                    invoice.invoice_date.strftime('%Y%m%d'),
                    # invoice.invoice_date.strftime('%d%m%Y'),
                    invoice.partner_shipping_id.name,
                    self.get_address(invoice.partner_shipping_id),
                    invoice.partner_shipping_id.city or '',
                    invoice.partner_shipping_id.country_id.name or '',
                    invoice.partner_shipping_id.x_school,
                ])

        return result

    def generate_xlsx_file(self, data_matrix):
        # Create a new workbook using XlsxWriter
        buffer = io.BytesIO()
        workbook = xlsxwriter.Workbook(buffer, {'in_memory': True})

        # Defining a bold format for the header
        bold_format = workbook.add_format({'bold': True})

        # Your data and formatting logic goes here
        worksheet = workbook.add_worksheet()

        # Seting the width of the columns
        # Headers are in the first row of data_matrix and their length determines the column width
        for col_num, header in enumerate(data_matrix[0]):
            column_width = 0
            if col_num in [6, 7]:
                column_width = len(header) + 10
            elif col_num == 4:
                column_width = len(header) + 20
            elif col_num == 5:
                column_width = len(header) + 30
            else:
                column_width = len(header) + len(header) * 0.5

            # Set the column width
            worksheet.set_column(col_num, col_num, column_width)

        # Write data to worksheet
        for row_num, row_data in enumerate(data_matrix):
            format_to_use = bold_format if row_num == 0 else None
            for col_num, cell_value in enumerate(row_data):
                worksheet.write(row_num, col_num, cell_value, format_to_use)

        # Close the workbook to save changes
        workbook.close()

        # Get the binary data from the BytesIO buffer
        binary_data = buffer.getvalue()
        return base64.b64encode(binary_data)

    def send_email(self, recipient_email, sender_email, subject, body, attachments, cc_email):
        mail_mail = self.env['mail.mail'].create({
            'email_from': 'your_email@example.com',
            'email_to': recipient_email,
            'email_from': sender_email,
            'email_cc': cc_email,
            'subject': subject,
            'body_html': body,
            'attachment_ids': [(0, 0, {'name': attachment[0], 'datas': attachment[1]}) for attachment in attachments],
        })
        mail_mail.send()

    def get_email_body(self):
        HEADER_TEXT = 'Weekly Apple Sales Report'
        table_width = 600

        email_content = {
            'text_line_1': 'Hi,',
            'text_line_2': f'Please find attached a {HEADER_TEXT}.',
            'text_line_3': 'Kind regards,',
            'text_line_4': 'JTRS Odoo',
            'table_width': table_width
        }

        email_html = f"""
        <!--?xml version="1.0"?-->
        <div style="background:#F0F0F0;color:#515166;padding:10px 0px;font-family:Arial,Helvetica,sans-serif;font-size:12px;">
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:5px auto;">
                <tbody>
                    <tr>
                        <td style="padding:0px;">
                            <a href="/" style="text-decoration-skip:objects;color:rgb(33, 183, 153);">
                                <img src="/web/binary/company_logo" style="border:0px;vertical-align: baseline; max-width: 100px; width: auto; height: auto;" class="o_we_selected_image" data-original-title="" title="" aria-describedby="tooltip935335">
                            </a>
                        </td>
                        <td style="padding:0px;text-align:right;vertical-align:middle;">&nbsp;</td>
                    </tr>
                </tbody>
            </table>
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:0px auto;background:white;border:1px solid #e1e1e1;">
                <tbody>
                    <tr>
                        <td style="padding:15px 20px 10px 20px;">
                            <p>{email_content['text_line_1']}</p>
                            </br>
                            <p>{email_content['text_line_2']}</p>
                            </br>
                            <p style="padding-top:20px;">{email_content['text_line_3']}</p>
                            <p>{email_content['text_line_4']}</p>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:15px 20px 10px 20px;">
                            <!--% include_table %-->
                        </td>
                    </tr>
                </tbody>
            </table>
            <table style="background-color:transparent;width:{email_content['table_width']}px;margin:auto;text-align:center;font-size:12px;">
                <tbody>
                    <tr>
                        <td style="padding-top:10px;color:#afafaf;">
                            <!-- Additional content can go here -->
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        """
        return email_html

    def generate_and_send_xlsx_file(self, recipient_email, sender_email, cc_email, data_matrix):
        HEADER_TEXT = 'Weekly Apple Sales Report'
        # sender_email = '"OdooBot" <odoobot@jtrs.co.uk>'
        # cc_email = 'martynas.minskis@jtrs.co.uk'

        # Generate XLSX file
        binary_data = self.generate_xlsx_file(data_matrix)

        # Define email parameters
        subject = f"{HEADER_TEXT} ({date.today().strftime('%d/%m/%y')})"
        body = self.get_email_body()

        # Using regular expression to replace '(', ')' and ' ' with '_'.
        attachment_name = re.sub(r'[() /]', '_', f"{subject}.xlsx")
        attachments = [(attachment_name, binary_data)]

        # Send email with the XLSX file attached
        self.send_email(recipient_email, sender_email,
                        subject, body, attachments, cc_email)

        return True

    def send_weekly_apple_sales_report(self, recipient_email, sender_email, cc_email):

        # Example usage:
        # data_matrix = [['A1', 'B1', 'C1'], ['A2', 'B2', 'C2'], ['A3', 'B3', 'C3']]
        data_matrix = self.get_and_process_invoice_data()
        if not data_matrix:
            _logger.warning('WARNING: data_matrix was not created.')
            return

        self.generate_and_send_xlsx_file(
            recipient_email, sender_email, cc_email, data_matrix)
