import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io


def find_text(element, tag, default=''):
    if element is None: return default
    found = element.find(tag)
    return found.text.strip() if found is not None and found.text is not None else default

def collect_all_fields(element, prefix=''):
    data = {}
    if element is None: return data
    for child in element:
        if '}' not in child.tag and len(list(child)) == 0:
            data[f"{prefix}{child.tag}"] = child.text.strip() if child.text else ''
    return data
    
def collect_extra_fields(element, prefix=''):
    data = {}
    if element is None: return data
    ttkhac = element.find('TTKhac')
    if ttkhac is not None:
        for ttin in ttkhac.findall('TTin'):
            truong = find_text(ttin, 'TTruong')
            dlieu = find_text(ttin, 'DLieu')
            if truong:
                clean_truong = ''.join(e for e in truong if e.isalnum())
                data[f"{prefix}Extra_{clean_truong}"] = dlieu
    return data

def flatten_tax_summary(ttoan_element, prefix='TToan_'):
    data = {}
    if ttoan_element is None: return data
    thtt_ltsuat = ttoan_element.find('THTTLTSuat')
    if thtt_ltsuat is not None:
        for ltsuat in thtt_ltsuat.findall('LTSuat'):
            tax_rate = find_text(ltsuat, 'TSuat')
            if tax_rate:
                thanh_tien = find_text(ltsuat, 'ThTien')
                tien_thue = find_text(ltsuat, 'TThue')
                data[f"{prefix}ThanhTien_{tax_rate}"] = thanh_tien
                data[f"{prefix}TienThue_{tax_rate}"] = tien_thue
    return data

def post_process_product_details(product_details):

    if 'THHDVu' in product_details:
        thhdvu_content = product_details['THHDVu']
        if ',' in thhdvu_content:
            parts = [p.strip() for p in thhdvu_content.split(',')]
            if len(parts) == 3:
                product_details['THHDVu'] = parts[0]      
                product_details['QuyCach'] = parts[1]    
                product_details['MHHDVu_Extracted'] = parts[2] 
    return product_details


def process_xml_to_excel_bytes_smarter(xml_file_content):
    try:
        root = ET.fromstring(xml_file_content)
    except ET.ParseError as e:
        st.error(f"Lỗi phân tích XML: {e}.")
        return None

    dlhdon = root.find('.//DLHDon')
    if dlhdon is None: st.error("Không tìm thấy thẻ <DLHDon>."); return None
    ndhdon = dlhdon.find('NDHDon')
    if ndhdon is None: st.error("Không tìm thấy thẻ <NDHDon>."); return None

    tt_chung, nban, nmua, ttoan = dlhdon.find('TTChung'), ndhdon.find('NBan'), ndhdon.find('NMua'), ndhdon.find('TToan')
    
    general_info = {}
    general_info.update(collect_all_fields(tt_chung, 'TTChung_'))
    general_info.update(collect_extra_fields(tt_chung, 'TTChung_'))
    general_info.update(collect_all_fields(nban, 'NBan_'))
    general_info.update(collect_extra_fields(nban, 'NBan_'))
    general_info.update(collect_all_fields(nmua, 'NMua_'))
    general_info.update(collect_extra_fields(nmua, 'NMua_'))
    general_info.update(collect_all_fields(ttoan, 'TToan_'))
    general_info.update(collect_extra_fields(ttoan, 'TToan_'))
    general_info.update(collect_extra_fields(dlhdon, 'DLHDon_'))
    general_info.update(collect_all_fields(dlhdon, 'DLHDon_'))
    general_info.update(flatten_tax_summary(ttoan, 'TToan_'))

    products_list = []
    dshhdvu = ndhdon.find('DSHHDVu')
    if dshhdvu is None: st.warning("Không tìm thấy danh sách hàng hóa (DSHHDVu)."); return None
        
    for hhdvu in dshhdvu.findall('HHDVu'):
        product_details = {**collect_all_fields(hhdvu), **collect_extra_fields(hhdvu)}

        product_details = post_process_product_details(product_details)
        
        products_list.append({**general_info, **product_details})

    if not products_list: st.warning("Không tìm thấy sản phẩm nào."); return None

    df_products = pd.DataFrame.from_records(products_list)
    df_general = pd.DataFrame([general_info])
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_products.to_excel(writer, sheet_name='ChiTietSanPham', index=False)
        df_general.to_excel(writer, sheet_name='ThongTinChung', index=False)
    
    return output.getvalue()

st.set_page_config(page_title="Trích xuất XML Hóa Đơn", layout="wide")
st.title("Trích xuất XML Hóa Đơn sang Excel")

uploaded_files = st.file_uploader("Chọn file XML Hóa đơn", type=['xml'], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        st.markdown(f"---")
        st.write(f"**Đang xử lý file: `{file.name}`**")
        
        xml_content = file.getvalue()
        excel_data = process_xml_to_excel_bytes_smarter(xml_content)
        
        if excel_data:
            st.success(f"Thành công `{file.name}`!")
            new_filename = file.name.replace('.xml', '.xlsx').replace('.XML', '.xlsx')
            st.download_button(
                label=f"Tải xuống: {new_filename}", data=excel_data,
                file_name=new_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_button_{file.name}"
            )
