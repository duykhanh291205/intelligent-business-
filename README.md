# SalesPulse — dashboard dự báo bán hàng linh hoạt

SalesPulse phân tích và dự báo dữ liệu bán hàng do người dùng tải lên. Ứng dụng không tự nạp bộ dữ liệu mẫu cố định. Giao diện xây bằng Streamlit; các mô hình dự báo dùng Python cơ bản để chạy trong môi trường hiện tại mà không cần NumPy/pandas/scikit-learn.

## Nạp dữ liệu

1. Tải tệp CSV, XLSX hoặc ZIP lên thanh bên. ZIP lồng ZIP (như bộ Walmart Kaggle) cũng được mở để chọn bảng.
2. Nếu tệp có nhiều bảng, chọn bảng lịch sử cần phân tích. Ứng dụng ưu tiên gợi ý `train.csv` và tránh chọn `test.csv`/`sample_submission.csv` làm dữ liệu huấn luyện.
3. Chọn cột ngày và cột số cần dự báo. Dashboard tự gợi ý cột theo tên; có thể đổi nếu gợi ý chưa đúng.
4. Tùy chọn cột số lượng, nhóm (khu vực/cửa hàng/danh mục) và mã đơn hàng.
5. Chọn mức tổng hợp ngày/tuần/tháng, thời hạn dự báo và mô hình.

Tối thiểu bảng cần có cột ngày và một cột số. Cột nhóm và số lượng không bắt buộc. App bỏ qua dòng thiếu ngày hoặc giá trị chỉ tiêu không phải số; sau đó tổng hợp các dòng theo ngày trước khi dự báo. Nếu dữ liệu đã ở dạng ma trận rộng (ví dụ mỗi ngày là một cột), chưa có cột ngày chuẩn, hoặc chỉ có dữ liệu `test` không có lịch sử thực tế thì cần chuyển/ghép bảng trước khi dự báo. Cần ít nhất 15 kỳ theo mức tổng hợp đã chọn.

Nếu bảng có đơn giá và số lượng nhưng chưa có doanh thu, chọn **Doanh thu (tự tính: đơn giá × số lượng)** trong phần ghép cột. Nếu đã tải nhiều bộ dữ liệu cùng lúc, chọn chế độ Rossmann hoặc bảng chung, rồi chọn đúng bảng cần phân tích.

Ví dụ để thử: tệp Superstore có `Order Date`, `Sales`, `Quantity`, `Region`; tệp Rossmann có `Date`, `Sales`, `Store`. Đây chỉ là ví dụ định dạng, người dùng có thể ghép các tên cột khác trong thanh bên.

### Rossmann (tùy chọn)

Khi tải `train.csv` Rossmann hoặc ZIP bộ dữ liệu, ứng dụng nhận diện cấu trúc và mở trang phân tích riêng theo cửa hàng. Có thể tải thêm `test.csv` và `store.csv` để dùng lịch tương lai và thông tin cửa hàng. Repository chỉ chứa mã nguồn, không đóng gói dữ liệu Kaggle; hãy tải dữ liệu riêng rồi chọn tải lên trong ứng dụng.

Với mọi tệp khác, ứng dụng dùng phần ánh xạ cột chung thay vì yêu cầu tên cột cố định.

## Chạy dự án trên Windows, macOS và Linux

Cài Python, mở terminal ngay tại thư mục dự án (thư mục có `app.py`), sau đó tạo môi trường ảo và cài thư viện cần thiết.

**Windows (PowerShell):**

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

**macOS hoặc Linux:**

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```

Các lệnh dùng đường dẫn tương đối trong thư mục dự án, không phụ thuộc vị trí cài Python trên máy. Streamlit sẽ in địa chỉ ứng dụng trong terminal, thường là `http://localhost:8501`. Dự án chỉ cần Streamlit; không yêu cầu NumPy/pandas/scikit-learn/XGBoost.

## Dashboard tương tác

Thanh điều hướng trên cùng chia nội dung thành các trang riêng để người dùng dễ tìm:

- **Tổng quan:** KPI trong kỳ, xu hướng doanh thu, cửa hàng nổi bật và so sánh promo.
- **Xu hướng:** chọn doanh thu/khách hàng, xem biểu đồ theo ngày/tuần/tháng và bảng số liệu.
- **Khám phá biểu đồ:** chọn kiểu đường, trung bình trượt, vùng, cột, thanh ngang, tròn, treemap, Pareto, tỷ trọng %, histogram, phân tán, heatmap hoặc bảng; biểu đồ thay đổi theo góc nhìn và dữ liệu.
- **Cửa hàng:** xếp hạng trong vùng lọc, xem chi tiết cửa hàng đã chọn và lịch sử theo thứ.
- **Khuyến mãi:** so sánh ngày promo, ngày lễ và ngày thường; đây là phân tích mô tả.
- **Dự báo:** biểu đồ thực tế/holdout/tương lai, so sánh mô hình, bảng dự báo và tải CSV.
- **Dữ liệu & phương pháp:** độ phủ dữ liệu, cách đánh giá và các giới hạn cần biết.

Với tệp bảng thông thường, thanh bên cho phép ghép cột, chọn cách tổng hợp, thời hạn và mô hình. Nếu tải bộ Rossmann, giao diện chuyên biệt có thêm bộ lọc bang/khu vực, thành phố, loại và mã cửa hàng; dự báo được tính trên lịch sử của cửa hàng đã chọn.

Dashboard dữ liệu chung có các trang Tổng quan, Xu hướng, Khám phá biểu đồ, Phân tích nhóm, Dự báo và Dữ liệu & phương pháp. Các trang cập nhật theo file và ánh xạ cột đã chọn.

## Phương pháp

- **Tệp bảng thông thường:** Linear Regression và Random Forest dự báo chuỗi chỉ tiêu đã tổng hợp theo ngày/tuần/tháng, sử dụng độ trễ và đặc trưng thời gian.
- **Rossmann:** ngoài hai mô hình trên, có Seasonal Baseline; Linear Regression và Random Forest sử dụng thêm trạng thái mở cửa, promo và ngày lễ.
- **Kiểm định:** giữ các kỳ cuối làm holdout theo thứ tự thời gian; không xáo trộn dữ liệu. Báo cáo WAPE, MAE và RMSE; Rossmann có thêm RMSPE.
- **Đánh giá trực quan:** trang Tổng quan tóm tắt xu hướng giữa hai nửa lịch sử, mức tập trung ở các nhóm dẫn đầu, so sánh dự báo với các kỳ gần đây và sai số holdout. Trong Khám phá biểu đồ, Pareto và tỷ trọng làm rõ đóng góp của từng nhóm; trung bình trượt giúp tách xu hướng khỏi dao động ngắn hạn.

## Nguồn, làm sạch và giới hạn

Rossmann Store Sales là cuộc thi Kaggle dự báo doanh thu (`Sales`) theo cửa hàng/ngày với thông tin khuyến mãi, ngày lễ, số khách và metadata cửa hàng. Mô hình hiện tại dự báo một cửa hàng được chọn trong giao diện, chưa tạo dự báo cho toàn bộ danh sách test để nộp Kaggle.

Dữ liệu train bị giới hạn trong khoảng thời gian của cuộc thi và có nhiều ngày cửa hàng đóng cửa. Sales âm bị chặn về 0; dự báo ngày đóng cửa được ép về 0 khi lịch tương lai có sẵn. Khi không có test.csv, app suy ra ngày cửa hàng mở từ lịch sử theo thứ trong tuần, đồng thời giả định không có promo/ngày lễ trong tương lai. Các kết quả vì thế phù hợp cho bài tập và so sánh mô hình, không phải kế hoạch vận hành thực tế.

Do Windows Application Control chặn NumPy trên máy hiện tại, app chưa dùng XGBoost hoặc thư viện scikit-learn. Có thể bổ sung XGBoost thật nếu môi trường sau này cho phép các thư viện đó.

## Gợi ý báo cáo/thuyết trình

1. Bài toán và quyết định kinh doanh: dự báo doanh thu từng cửa hàng 6 tuần.
2. Dataset, các trường dùng và quy tắc xử lý ngày đóng cửa/giá trị thiếu.
3. Đặc trưng trễ/mùa vụ và cách chia holdout theo thời gian.
4. Bảng so sánh baseline, Linear Regression, Random Forest qua RMSPE/WAPE/MAE/RMSE.
5. Demo bộ dữ liệu đã chọn, các biểu đồ và giới hạn của dự báo. Nếu dùng Rossmann, trình bày thêm phân tích cửa hàng và khuyến mãi.
