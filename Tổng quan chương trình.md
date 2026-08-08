# **Tổng quan về SignFill**

1. #### **chia ra 3 thư mục input, output, templates riêng để người dùng có thể dễ nhận biết.**
* ##### *thư mục input:*

\- chứa các file excel.

\- trong file chứa N cột với mỗi hàng là thông tin của 1 người (ví dụ các cột: tên, tuổi, CCCD, ngày cấp, ... và các hàng là: nguyễn văn A, trần văn B, ...).

\- người dùng có thể thêm hoặc bớt các cột thông tin khác, phần mềm tự nhận diện và điền vào template để xuất ra file output.

* ##### *thư mục templates:*

\- chứa các file word dạng template.

\- trong các template sẽ bao gồm các trường do người dùng tự thêm hoặc bớt.

\- các trường đó có quy luật dễ nhớ, dễ dùng đối với người không biết lập trình.

\- các trường phải tự nhận diện và đồng bộ số lượng với các trường trong file excel (phải đảm bảo đúng thông tin từng trường).

\- quy tắc duy nhất người dùng cần nhớ: Tên trong {{...}} của Word phải giống tên cột Header trong Excel.

* ##### *thư mục output:*

\- xuất ra các file với định dạng docx và tên do người dùng có thể tùy chỉnh (ví dụ như: \[ngày tháng năm ký HĐ].\[số HĐ].HĐLĐ-\[tên công ty].\[tên người được ký HĐ] - tương ứng với: 260601.75.HĐLĐ-HĐP.Nguyễn Duy Hưng.docx) - các thông tin này sẽ lấy từ file excel (lưu ý: đây không phải mẫu tên cố định, người dùng có thể đặt các tên khác với thông tin được lấy từ file excel).



**\*\*\* Cấu trúc dự kiến: các trường trong template sẽ tương ứng với hàng tên của các cột trong excel mà người dùng tạo nếu làm thủ công (Có cách nào tự động không?) \*\*\***



#### **2. kiểu Generate: One Person → One Contract (100 persons → 100 contracts)**



#### **3. UI: phiên bản đầu chỉ cần** 

\- 1 button Generate (sau khi bấm thông báo thành công hoặc không thành công kèm với các lỗi cụ thể)

\- 1 phần để người dùng có thể tùy chỉnh tên của file output (nhiều file của nhiều người giống như phần thư mục output đã nêu ở trên)

#### 

