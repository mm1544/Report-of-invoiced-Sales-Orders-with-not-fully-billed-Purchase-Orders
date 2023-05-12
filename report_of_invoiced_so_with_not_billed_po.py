# This script finds invoiced Sale Orders with not fully billed Purchase Orders, generates HTML for email template and
# sends the email to designated person. Report covers one month period.

HEADER_TEXT = 'Report of invoiced Sales Orders with not fully billed Purchase Orders'
EMAIL_TO = 'martynas.minskis@gmail.com'
EMAIL_CC = 'martynas.minskis@gmail.com'
EMAIL_TEMPLATE_ID = 315
EMAIL_SERVER_ID = 10
TODAY_DATE = datetime.date.today()


def is_today_first_day_of_month():
  if TODAY_DATE.day == 1:
    return True


# Returns recordset of invoiced SOs created in previous month.
def find_invoiced_sale_orders():
  year = TODAY_DATE.year
  month = TODAY_DATE.month
  # If we have 01.01.2023, then we want the date boundary to be 01.12.2022
  if month == 1:
    year -= 1
    month = 12
  else:
    month -= 1
  
  time_boundary = datetime.datetime(year,month,1)
    
  sale_orders = env['sale.order'].search([('create_date','>',time_boundary),('invoice_status','in',['invoiced'])])
  
  return [sale_orders, time_boundary]
  

# Returns SO id with not billed POs ids in form {so_id_int : [po_id1_int, po_id2_int, ...], ...}
def get_sos_with_not_billed_pos(so_recordset):
  res = {}
  for so in so_recordset:
    not_billed_pos = env['purchase.order'].sudo().search([('x_sale_id', '=', so.id),('invoice_status', '!=', 'invoiced'),('state','in',['purchase','done'])])
    if not not_billed_pos:
      continue

    res[so.id] = [po.id for po in not_billed_pos]
  
  return res
  

# Create table base html
def create_base_table_template():
  style_1 = 'padding: 5px 10px; border:2px solid #E6E6E6; color:#515166;'
  
  t  ='<div style="text-align:center;">'
  t +='<!--% include_header %-->'
  t +='  <table style="background:#F0F0F0; color:#515166; font-family:Arial,Helvetica,sans-serif; font-size:12px; border-collapse:collapse; border:2px solid #E6E6E6; margin:0px auto;">'
  t +='      <thead style="% style_1 %">'
  t +='          <tr style="margin:0px; padding:0px;">'
  t +='              <th></th>'
  t +='              <th style=" white-space:nowrap; % style_1 %">Sales Orders</th>'
  t +='              <th style=" white-space:nowrap; % style_1 %">Purchase Orders</th>'
  t +='              <th style=" white-space:nowrap; % style_1 %">Customer</th>'
  t +='          </tr>'
  t +='      </thead>'
  t +='      <tbody>'
  t +='<!--% include_table_lines %-->'
  t +='      </tbody>'
  t +='  </table>'
  t +='</div>'
  
  t = t.replace('% style_1 %', style_1)
  return t
  
  
def get_pos_link_html(po_obj_list):
  p = ''
  separator = ', '
  if not po_obj_list:
    return
  
  for index, po_obj in enumerate(po_obj_list):
    if index == len(po_obj_list)-1:
      separator = ''
      
    p += '<a href="/web#model={}&amp;id={}" target="_blank" style="color:#008784; text-decoration: none;">{}</a>{}'.format('purchase.order', po_obj.id, po_obj.name, separator)
    
  return p
  
  
def create_table_row(row_data_dict):
  style_1 = 'text-align:left; border:2px solid #E6E6E6; padding:3px 10px;'
  if not row_data_dict:
    return


  line_number = row_data_dict['line_number']
  sales_order_obj = row_data_dict['sales_order_obj']
  purchase_orders_obj_list = row_data_dict['purchase_orders_obj_list']

  
  # Create POs link html
  pos_link_html = get_pos_link_html(purchase_orders_obj_list)
  
    
  r  ='<tr>'
  r +=' <td style="% style_1 %">{}</td>'.format(line_number)
  r +=' <td style="% style_1 %"><a href="/web#model={}&amp;id={}" target="_blank" style="color:#008784; text-decoration: none;">{}</a></td>'.format('sale.order', sales_order_obj.id, sales_order_obj.name)
  r +=' <td style="% style_1 %">{}</td>'.format(pos_link_html)
  r +=' <td style="% style_1 %"><a href="/web#model={}&amp;id={}" target="_blank" style="color:#008784; text-decoration: none;">{}</a></td>'.format('res.partner', sales_order_obj.partner_id.id, sales_order_obj.partner_id.name)
  r +='</tr>'
  
  r = r.replace('% style_1 %', style_1)
  return r
  
  
## Creates all table lines for report table
## @sos_and_pos_dict -> {so_id_int : [po_id1_int, po_id2_int, ...], ...}
def create_all_table_lines(sos_and_pos_dict):
  table_rows_html = ''
  line_count = 0
  
  if not sos_and_pos_dict:
    return ''
  
  so_id_ordered_list = list(sos_and_pos_dict.keys())
  so_id_ordered_list.sort()
  

  for so_id in so_id_ordered_list:
    line_count += 1
    

    data_dict = {
      'line_number': line_count,
      'sales_order_obj': env['sale.order'].browse([so_id]),
      'purchase_orders_obj_list': env['purchase.order'].browse(sos_and_pos_dict[so_id])
    }
    
    table_rows_html += create_table_row(data_dict)
  
  return table_rows_html
  
  
def create_report_table_html(data_dict):
  if not data_dict:
    return ''

  all_table_lines_html = create_all_table_lines(data_dict)
  
  if not all_table_lines_html:
    return
  
  table_base = create_base_table_template()

  full_table = table_base.replace('<!--% include_table_lines %-->', all_table_lines_html)

  return full_table
  
  
def create_report_email_html(report_time_boundary):
  
  text_line_1 = 'Hi,'
  text_line_2 = 'Please see {}.'.format(HEADER_TEXT)
  text_line_2_2 = 'Covered period of time from <strong>{}</strong>'.format(report_time_boundary.strftime('%d/%m/%Y'))
  text_line_4 = 'Kind regards,'
  text_line_5 = 'JTRS Odoo'
  
  e = '<!--?xml version="1.0"?-->'
  e += '<div style="background:#F0F0F0;color:#515166;padding:10px 0px;font-family:Arial,Helvetica,sans-serif;font-size:12px;">'
  e += '<table style="background-color:transparent;width:600px;margin:5px auto;">'
  e += '<tbody>'
  e += '<tr><td style="padding:0px;">'
  e += '<a href="/" style="text-decoration-skip:objects;color:rgb(33, 183, 153);"><img src="/web/binary/company_logo" style="border:0px;vertical-align: baseline; max-width: 100px; width: auto; height: auto;" class="o_we_selected_image" data-original-title="" title="" aria-describedby="tooltip935335"></a>'
  e += '</td><td style="padding:0px;text-align:right;vertical-align:middle;">&nbsp;</td></tr>'
  e += '</tbody>'
  e += '</table>'
  e += '<table style="background-color:transparent;width:600px;margin:0px auto;background:white;border:1px solid #e1e1e1;">'
  e += '<tbody>'
  e += '<tr><td style="padding:15px 20px 10px 20px;">'
  e += '<p style="">{}</p>'.format(text_line_1)
  e += '<br>{}'.format(text_line_2)
  e += '<br>{}'.format(text_line_2_2)
  e += '<br><br>'
  e += '<span>'
  e += '      <!--% include_table %-->'
  e += '</span>'
  e += '<br><br><br>'
  e += '<span style="margin-left:30px;font-weight:normal;">{}</span>'.format(text_line_4)
  e += '<br><span style="margin-left:30px;font-weight:normal;"><br></span><span style="margin-left:30px;font-weight:normal;">{}</span>'.format(text_line_5)
  e += '</td></tr>'
  e += '<tr>'
  e += '</tr>'
  e += '</tbody>'
  e += '</table>'
  e += '<table style="background-color:transparent;width:600px;margin:auto;text-align:center;font-size:12px;">'
  e += '<tbody>'
  e += '<tr><td style="padding-top:10px;color:#afafaf;">'
  e += '<p style=""></p>'
  e += '<p style=""></p>'
  e += '</td></tr>'
  e += '</tbody>'
  e += '</table>'
  e += '</div>'
  return e
  
  
def logging():
  mess = "{} was sent".format(HEADER_TEXT)
  log(mess, level='info')


def send_email(report_html):

  email_template = env['mail.template'].browse([EMAIL_TEMPLATE_ID])
  if not email_template:
    return
        
  email_subject = '{} ({})'.format(HEADER_TEXT, datetime.datetime.now().strftime('%d/%m/%y'))
  
  email_template.write({
    'body_html': report_html,
    'subject': email_subject,
    'email_from': '"OdooBot" <odoobot@jtrs.co.uk>',
    'email_to': EMAIL_TO,
    'email_cc': EMAIL_CC,
    'mail_server_id': EMAIL_SERVER_ID
  })
  
  partner_obj = env['res.partner'].browse(1)
  if not partner_obj:
    return
  
  ### Sending email
  email_template.send_mail(partner_obj.id,force_send=True)
  
  logging()
      
    
  
def main():
  if not is_today_first_day_of_month():
    return
  allinvoiced_sos, time_boundary = find_invoiced_sale_orders()
  
  if not allinvoiced_sos:
    mess = 'No Sales Orders found (allinvoiced_sos)'
    log(mess, level='info')
    return
  
  sos_with_not_billed_pos_dict = get_sos_with_not_billed_pos(allinvoiced_sos)
  
  report_table_html = create_report_table_html(sos_with_not_billed_pos_dict)
  report_email_html = create_report_email_html(time_boundary)
  
  email_html_with_table = report_email_html.replace('<!--% include_table %-->', ('<br>' + report_table_html))

  send_email(email_html_with_table)
    

try:
  main()
except Exception as e:
  mess = 'Error: {}'.format(e)
  log(mess, level='error')
  