import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io


def find_text(element, tag, default=''):
    if element is None:
        return default
    found = element.find(tag)
    return found.text.strip() if found is not None and found.text is not None else default

def collect_all_fields(element, prefix=''):
    data = {}
    if element is None:
        return data
    for child in element:

        if '}' not in child.tag and len(list(child)) == 0:
            field_name = f"{prefix}{child.tag}"
            data[field_name] = child.text.strip() if child.text else ''
    return data
    
def collect_extra_fields(element, prefix=''):
    data = {}
    if element is None:
        return data
    ttkhac = element.find('.//TTKhac')
    if ttkhac is not None:
        for ttin in ttkhac.findall('TTin'):
            truong = find_text(ttin, 'TTruong')
            dlieu = find_text(ttin, 'DLieu')
            if truong:
                field_name = f"{prefix}Extra_{truong}"
                data[field_name] = dlieu
    return data

def process_xml_to_excel_bytes_smarter(xml_file_content):
    try:
        root = ET.fromstring(xml_file_content)
    except ET.ParseError as e:
        st.error(f"Lỗi phân tích XML: {e}.")
        return None

    tt_chung = root.find('.//TTChung')
    nban = root.find('.//NBan')
    nmua = root.find('.//NMua')
    ttoan = root.find('.//TToan')
    
    general_info = {}
    general_info.update(collect_all_fields(tt_chung, 'TTChung_'))
    general_info.update(collect_extra_fields(tt_chung, 'TTChung_'))
    general_info.update(collect_all_fields(nban, 'NBan_'))
    general_info.update(collect_extra_fields(nban, 'NBan_'))
    general_info.update(collect_all_fields(nmua, 'NMua_'))
    general_info.update(collect_extra_fields(nmua, 'NMua_'))
    general_info.update(collect_all_fields(ttoan, 'TToan_'))

    products_list = []
    dshhdvu = root.find('.//DSHHDVu')
    if dshhdvu is None:
        st.warning("Không tìm thấy danh sách hàng hóa dịch vụ (DSHHDVu) trong file.")
        return None
        
    for hhdvu in dshhdvu.findall('HHDVu'):
        product_details = {}

        product_details.update(collect_all_fields(hhdvu))

        product_details.update(collect_extra_fields(hhdvu))

        product_row = {**general_info, **product_details}
        products_list.append(product_row)

    if not products_list:
        st.warning("Không tìm thấy sản phẩm nào trong danh sách.")
        return None


    df_products = pd.DataFrame.from_records(products_list)
    df_general = pd.DataFrame([general_info])
    

    cols_order = sorted(df_general.columns)
    df_general = df_general[cols_order]
    
    product_cols_order = cols_order + sorted([col for col in df_products.columns if col not in cols_order])
    df_products = df_products[product_cols_order]


    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_products.to_excel(writer, sheet_name='ChiTietSanPham', index=False)
        df_general.to_excel(writer, sheet_name='ThongTinChung', index=False)
    
    processed_data = output.getvalue()
    return processed_data

st.set_page_config(page_title="Trích xuất XML Hóa Đơn", layout="wide")
st.title("Trích xuất XML Hóa Đơn sang Excel")

uploaded_file = st.file_uploader(
    "Chọn file XML Hóa đơn",
    type=['xml'],
    accept_multiple_files=True
)

if uploaded_file:
    for file in uploaded_file:
        st.markdown(f"---")
        st.write(f"**Đang xử lý file: `{file.name}`**")
        
        xml_content = file.getvalue()
        excel_data = process_xml_to_excel_bytes_smarter(xml_content)
        
        if excel_data:
            st.success(f"Thành công `{file.name}`!")
            
            new_filename = file.name.replace('.xml', '.xlsx').replace('.XML', '.xlsx')

            st.download_button(
                label=f"Tải xuống: {new_filename}",
                data=excel_data,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            )
