import streamlit as st
import os
import pandas as pd
from PIL import Image
from helper_func import * 

SUPPORTED_RAG_FILE_TYPES = [
    # Tài liệu văn phòng
    "pdf",
    "doc",
    "docx",
    "pptx",
]

SUPPORTED_PANDAS_FILE_TYPES = [
    "csv",        # Dữ liệu bảng dạng text, phổ biến nhất
    "xlsx",       # Excel (hỗ trợ nhiều sheet)
    "xls",        # Excel cũ
]


st.set_page_config(page_title="Assistant", page_icon=":robot_face:", layout='wide')

st.sidebar.header(':robot_face: May I help you ?')

choose = st.sidebar.radio(
    "Choose features",
    ["Chatbot RAG", "Report Generator"],
    captions=[
        "Answer user questions based on the content of the document you uploaded",
        "Automatically analyze, synthesize, and visualize data."
    ],
)
save_option = {
    'statistics': False,
    'variables': [],
    'correlations': False
}
if choose == 'Chatbot RAG':
    st.write(check_api_key())
    
    # Tải file lên
    files = st.sidebar.file_uploader(":file_folder: Upload a file", type=SUPPORTED_RAG_FILE_TYPES, accept_multiple_files=True)
    # Kiểm tra session_state
    if 'chat_history' not in st.session_state:
            st.session_state.chat_history = load_history_chat()
    # Kiểm tra trùng file
    uploaded_names = [file.name for file in files]
    if len(uploaded_names) != len(set(uploaded_names)):
        st.sidebar.warning("You have uploaded files with the same name!")
        # Bỏ file bị trùng để giảm dung lượng bộ nhớ
        unique_files = []
        seen_names = set()
        for file in files:
            if file.name not in seen_names:
                unique_files.append(file)
                seen_names.add(file.name)
        files = unique_files
    if files:
        # Tách nội dụng từ trong các files
        raw_text = get_file_text(files)
        if raw_text:
            # Chia nhỏ nội dung thành các đoạn nhỏ trong mảng list
            text_chunks = get_text_chunk(raw_text)
            if text_chunks:
                # Chuyển hoá văn bản thành từ khoá để hỗ trợ tìm kiếm
                get_vector_store(text_chunks)
            else:
                st.error('Check the document content again.')
        else:
            st.error('Không đọc được nội dung tài liệu.')
        # Thông báo khi file tải lên thành công
        current_file_names = [file.name for file in files]
        if 'uploaded_file_names' not in st.session_state:
            st.session_state.uploaded_file_names = []
        else:
            if current_file_names < st.session_state.uploaded_file_names:
                st.session_state.uploaded_file_names = current_file_names
        # Nếu danh sách file thay đổi thì assistant mới gửi thông báo
        if current_file_names != st.session_state.uploaded_file_names and len(files) > 0:
            assistant_msg = f"Uploaded successfully {len(files)} file: " + ", ".join(current_file_names)
            st.session_state.chat_history.append({'role': 'assistant', 'content': assistant_msg})
            save_chat_history()
            st.session_state.uploaded_file_names = current_file_names
        # Tiêu đề
        st.title(":mag_right: Chatbot RAG")
        # Khung chat của user
        user_question = st.chat_input('Please ask after the document has been analyzed.')
        if user_question:
            st.session_state.chat_history.append({'role': 'user', 'content': user_question})
            response = user_input(user_question)
            st.session_state.chat_history.append({'role': 'assistant', 'content': response})
            save_chat_history()
        # khởi tạo session state để lưu lại lịch sử
        if 'chat_history' not in st.session_state:
            st.session_state.chat_history = load_history_chat()
        # hiển thị lịch sử chat
        for msg in st.session_state.chat_history:
            with st.chat_message(msg['role']):
                    st.markdown(msg['content'])
        # tạo button xoá lịch sử chat
        if len(st.session_state.chat_history):
            if st.button('Xoá lịch sử chat'):
                st.session_state.chat_history = []
                st.session_state.uploaded_file_names = []  # Xoá luôn danh sách tên file đã upload
                if os.path.exists('chat_history.json'):
                    os.remove('chat_history.json')
                st.rerun()
       
    else: 
        st.title("📥 Please share me your data.")
        if os.path.exists('chat_history.json'):
            os.remove('chat_history.json')
elif choose == 'Report Generator':
    # Tiêu đề
    st.title(":bar_chart: Report Generator")
     # Tải file lên
    file = st.sidebar.file_uploader(":file_folder: Upload a file", type=SUPPORTED_PANDAS_FILE_TYPES)
    
    if "reports" not in st.session_state:
        st.session_state.reports = []

    if "no" not in st.session_state:
        st.session_state.no = 0

    element_num = 0

    chart_folder_name = "charts"
    chart_folder_path = f"./{chart_folder_name}"

    if not os.path.exists(chart_folder_name):
        os.makedirs(chart_folder_name)

    if file:
        # Kiểm tra xem đuôi file
        if file.name.lower().endswith('.csv'):
            # Xử lý file csv
            data = pd.read_csv(file)
        else:
            # xử lý file excel
            data = pd.read_excel(file)
        # Tách các cột
        object_columns = data.select_dtypes(include="object").columns
        data[object_columns] = data[object_columns].astype("string")
        # Categorical values
        categorical_columns = data.select_dtypes(include=["string"]).columns.tolist()
        # Numerical values
        numeric_columns = data.select_dtypes(include=["float64", "int64"]).columns.tolist()
        numeric_columns_1 = data.select_dtypes(include=["float64", "int64"]).columns.tolist()
        numeric_columns_2 = data.select_dtypes(include=["float64", "int64"]).columns.tolist()
        # Thêm selectbox
        choose_overview = st.selectbox('Please select: ', ('Dataset statistics', 'Variables',  'Correlations'))
        # thêm overview
        if choose_overview == 'Dataset statistics':
            st.subheader("Dataset statistics")
            overview_dict = {
                'Number of variables': data.shape[1],
                'Number of observations': data.shape[0],
                'Missing cells': data.isnull().sum().sum(),
                'Missing cells (%)': round(data.isnull().sum().sum() / (data.shape[0]*data.shape[1]) * 100, 2) if data.shape[0]*data.shape[1] > 0 else 0,
                'Duplicate rows': data.duplicated().sum(),
                'Duplicate rows (%)': round(data.duplicated().sum() / data.shape[0] * 100, 2) if data.shape[0] > 0 else 0,
            }
            data_overview = pd.DataFrame(list(overview_dict.items()), columns=['Metric', 'Value'])
            data_overview['Value'] = data_overview['Value'].apply(lambda x: int(x) if isinstance(x, (int, float)) and x == int(x) else x)
            st.table(data_overview)
           
        elif choose_overview == 'Variables':
            st.subheader("Variables")
            choose_feature = st.selectbox('Please select: ', list(data.columns))
            overview_dict = {
                'Distinct': data[choose_feature].nunique(),
                'Missing cells': data[choose_feature].isnull().sum().sum(),
                'Missing cells (%)': round(data[choose_feature].isnull().sum() / data.shape[1] * 100, 2),
            }
            data_overview = pd.DataFrame(list(overview_dict.items()), columns=['Metric', 'Value'])
            data_overview['Value'] = data_overview['Value'].apply(lambda x: int(x) if isinstance(x, (int, float)) and x == int(x) else x)
            st.table(data_overview)
           
        elif choose_overview == 'Correlations':
            st.subheader("Correlations")
            corr = data.corr(numeric_only=True)
            st.dataframe(corr.style.background_gradient(cmap='viridis'))
        st.subheader("Save data into excel")    
        agree_1 = st.checkbox("Dataset statistics")
        if agree_1:
            save_option['statistics'] = True
        options = st.multiselect(
            "Variables",
            list(data.columns),
            default=[],
        )
        if options:
            save_option['variables'] = options
        agree_2 = st.checkbox("Correlations")
        if agree_2:
            save_option['statistics'] = True
        st.subheader("Aggregation Options")
        # thêm radio
        genre = st.radio(
            "What's your favorite movie genre",
            ["Aggregation", "Interactions"],
        )
            
        if len(categorical_columns) > 0 and len(numeric_columns) > 0:
            #set session_state
            if "category" not in st.session_state:
                st.session_state.category = categorical_columns[0]

            if "numeric" not in st.session_state:
                st.session_state.numeric = numeric_columns[0]
            if "numeric1" not in st.session_state:
                st.session_state.numeric1 = numeric_columns_1[0]
            if "numeric2" not in st.session_state:
                st.session_state.numeric2 = numeric_columns_2[0]
            # chọn cách vẽ bản đồ
            if genre == "Aggregation":
                category_col = st.selectbox("Choose a categorical column for grouping:", categorical_columns, key="category")
                numeric_col = st.selectbox("Choose a numeric column for aggregation:", numeric_columns, key="numeric")
                aggregation_function = st.selectbox(
                    "Aggregation function: ",
                    ["sum", "mean", "count", "min", "max"]
                )

                if aggregation_function == "sum":
                    aggregated_data = data.groupby(category_col)[numeric_col].sum().reset_index()
                elif aggregation_function == "mean":
                    aggregated_data = data.groupby(category_col)[numeric_col].mean().reset_index()
                elif aggregation_function == "count":
                    aggregated_data = data.groupby(category_col)[numeric_col].count().reset_index()
                elif aggregation_function == "min":
                    aggregated_data = data.groupby(category_col)[numeric_col].min().reset_index()
                elif aggregation_function == "max":
                    aggregated_data = data.groupby(category_col)[numeric_col].max().reset_index()

                # if aggregated_data is not None:
                st.subheader(f"Aggregated Data: {aggregation_function} of {numeric_col} by {category_col}")
                # Tạo bảng màu hệ viridis
                st.dataframe(aggregated_data.style.background_gradient(cmap='viridis'))
                # Vẽ bản đồ
                st.subheader("Visualize your data")
                chart_type = st.selectbox(
                    "Choose chart type",
                    ["Line Chart", "Bar Chart", "Scatter Plot", "Pie Chart"]
                )
            else:
                category_col = st.selectbox("Choose a numeric column for aggregation 1:", numeric_columns_1, key="numeric1")
                numeric_col = st.selectbox("Choose a numeric column for aggregation 2:", numeric_columns_2, key="numeric2", index=1)
                aggregated_data = data[[category_col, numeric_col]]
            if st.button("Plot Graph"):
                if genre == "Aggregation":
                    chart_path, chart_name = plot_chart(chart_folder_path, chart_type, aggregated_data,st.session_state.category,st.session_state.numeric)
                else:
                    chart_path, chart_name = plot_chart_interactions(chart_folder_path, aggregated_data,st.session_state.numeric1,st.session_state.numeric2)
                
                response = generate_report_from_chart(chart_folder_path, chart_name)
                
                report = {
                    'pivot_table': aggregated_data,
                    'chart_path': chart_path,
                    'sheet_name': f'Sheet {st.session_state.no}',
                    'insight': response
                }
                
                st.session_state.reports.append(report)
                st.session_state.no += 1
                
            # Hiển thị các chart ở đây
            if st.session_state.reports != []:
                for i, report in enumerate(st.session_state.reports):
                    filename = report["chart_path"]
                    if filename.endswith(('.png', '.jpg', '.jpeg')):
                        file_path = os.path.join(chart_folder_path, filename)

                        if os.path.exists(file_path):
                            image = Image.open(file_path)
                            st.image(image, caption=filename, use_column_width=True)
                        else:
                            st.warning(f"File {file_path} không tồn tại!")
                        element_num += 1
                        if st.button("Remove chart", key=element_num):
                            remove_chart(file_path)
                            st.session_state.reports.pop(i)
                            st.rerun()
                        
                        
                        
            if st.button("Generate Report"):
                generate_excel_report(data, st.session_state.reports, 'report', save_option)
                with open("report.xlsx", "rb") as file:
                    excel_data = file.read()
                st.download_button(label="Download", data=excel_data, file_name="report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        