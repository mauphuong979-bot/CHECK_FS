"""Hộp thoại "Hướng dẫn sử dụng" mở từ nút ❓ trên thanh trên cùng."""

from __future__ import annotations

import streamlit as st

QUICK_START = """
**Ứng dụng hỗ trợ kiểm toán viên kiểm tra tính chính xác số học và tính nhất quán của số liệu trên
Báo cáo tài chính (BCTC).** Hiện chỉ nhận tệp **Word .docx**; chưa hỗ trợ .doc, .pdf, .xlsx.

| Bước | Việc cần làm | Ở đâu |
|---|---|---|
| 1 | Tải BCTC dạng .docx cần kiểm tra | **📂 Báo cáo cần kiểm tra** |
| 2 | Bấm **Tạo báo cáo kiểm tra** và chờ ứng dụng đọc bảng, nhận diện Thông tư | Nút bên dưới vùng tải tệp |
| 3 | Đọc kết luận và danh sách sai lệch: mỗi dòng có chỉ tiêu, bảng (BS/PL/CF/TM, đơn vị tiền tệ), số trình bày, số tính lại và chênh lệch | Khung kết quả |
| 4 | Tải **Word có comment**; tải **Excel chi tiết** khi cần truy vết công thức | Nút tải trong khung kết quả |
| 5 | Rà soát comment trong Word và sheet chi tiết trong Excel | Microsoft Word / Excel |

**Các nút ở góc phải trên cùng**
* **🌐 Đối chiếu số với bản dịch** – mở ứng dụng so khớp số giữa BCTC tiếng Việt và bản dịch.
* **❓** – mở hướng dẫn này.
* **👤 Tên tài khoản** – đổi mật khẩu và **Đăng xuất**.
"""

PROCESS = """
**1. Chuẩn bị tệp nguồn**
* Dùng bản BCTC Word .docx rõ ràng, số liệu nằm trong bảng; tệp không bị khóa (Protected/Read-only).

**2. Tải tệp và tạo báo cáo**
* Chọn tệp rồi bấm **Tạo báo cáo kiểm tra**. Ứng dụng không tự xử lý khi vừa tải tệp.
* Khi chọn tệp khác, kết quả cũ được ẩn cho đến khi bấm tạo báo cáo cho tệp mới.

**3. Nhận diện Thông tư**
* Ứng dụng nhận diện chế độ kế toán theo viện dẫn Thông tư trong báo cáo.
* Hiện chỉ có bộ rule nghiệp vụ đã xác nhận cho **Thông tư 200/2014/TT-BTC**. Với Thông tư 133 hoặc
  Thông tư 99 ứng dụng chỉ cảnh báo, **không** tự áp rule của Thông tư 200.

**4. Phạm vi kiểm tra tự động**
* Cộng dọc, cộng ngang, công thức theo mã chỉ tiêu, dấu của một số chỉ tiêu Lưu chuyển tiền tệ, quan hệ
  giữa các kỳ và đối chiếu Thuyết minh với Bảng cân đối kế toán, Kết quả kinh doanh, Lưu chuyển tiền tệ.
* So sánh chính xác tuyệt đối, không làm tròn và không có dung sai: chênh lệch 0,01 vẫn là sai lệch.
* Bảng thông tin (mục lục, chữ ký, hợp đồng vay, thời gian khấu hao…) được ghi **Không kiểm tra**.

**5. Đọc khung kết quả**
* Dòng đầu là kết luận: số **sai lệch** và số **cần xem xét**; dòng nhỏ bên dưới cho biết đã kiểm tra bao nhiêu bảng.
* Danh sách **Sai lệch** hiển thị sẵn, chênh lệch ở cột bên phải. Nhóm **Cần xem xét** thu gọn, bấm để mở.
* Căn cứ nhận diện Thông tư nằm trong mục **Chi tiết nhận diện Thông tư và phạm vi kiểm tra**.
"""

READ_EXCEL = """
**Sheet `00_Tong_hop`**
* Mỗi dòng là một bảng trong BCTC: **STT bảng**, **Sheet chi tiết** (có liên kết), **Nội dung bảng**,
  **Loại bảng**, **Trạng thái**, **Tổng sai sót**, **Trạng thái đối chiếu BCTC**, **Ghi chú**, **Số phép kiểm tra**.
* Trạng thái đối chiếu: `Khớp`, `Có chênh lệch`, `Không tìm thấy số liệu TM`, `Không tìm thấy chỉ tiêu BCTC`,
  `Không áp dụng`.

**Màu tab sheet**
* **Đỏ sẫm** – bảng có sai lệch.
* **Vàng** – bảng cần kiểm toán viên xem xét.
* **Xanh lá** – bảng đã kiểm tra và khớp.

**Sheet chi tiết**
* Dòng 1 có thanh điều hướng **Trở về Tổng hợp**, **Trước**, **Tiếp**; có thể dùng thêm
  `Ctrl + Page Up/Down` để chuyển sheet.
* Ô công thức kiểm tra nền xám nhạt; dữ liệu cần xác minh nền vàng.
* Sheet Thuyết minh có khối đối chiếu với BCTC ngay dưới dữ liệu, căn đúng cột kỳ của Thuyết minh.
"""

READ_WORD = """
**File Word có comment**
* Comment được neo vào đúng ô trong bảng của báo cáo gốc; ứng dụng **không sửa** số liệu hay nội dung báo cáo.
* Mỗi comment ghi rõ số trình bày, số tính lại và chênh lệch; nội dung bắt đầu bằng `[SAI LỆCH]` hoặc
  `[CẦN XEM XÉT]`.
* Lọc theo người comment **CHECK_FS - Sai lệch** hoặc **CHECK_FS - Cần xem xét** tại
  **Review → Show Markup → Specific People**.

**Phân biệt mức độ**
* **🔴 Sai lệch** – lỗi số học hoặc đối chiếu không khớp xác định được.
* **🟡 Cần xem xét** – chưa đủ căn cứ để tự động kết luận (thiếu dữ liệu, ghép suy luận, cấu trúc bảng
  chưa nhận diện…); kiểm toán viên cần tự xác minh.
* Kết quả đã khớp chỉ ghi trong Excel, không tạo comment Word.
"""

ACCOUNT = """
**Đăng nhập**
* Chọn **Tên đăng nhập**, nhập **Mật khẩu**, bấm **Đăng nhập**.
* **Ghi nhớ đăng nhập trên máy này**: không cần nhập lại mật khẩu trong thời gian ghi nhớ trên trình duyệt đó.
  Không dùng trên máy dùng chung.
* Nhập sai mật khẩu nhiều lần liên tiếp sẽ phải chờ một lúc trước khi thử lại.

**Đổi mật khẩu**
* Bấm **👤 Tên tài khoản → 🔑 Đổi mật khẩu**, nhập mật khẩu hiện tại và mật khẩu mới (tối thiểu 6 ký tự).
* Sau khi đổi, các trình duyệt khác đã ghi nhớ đăng nhập sẽ phải đăng nhập lại.

**Đăng xuất**
* **👤 Tên tài khoản → 🚪 Đăng xuất** xóa ghi nhớ trên trình duyệt này và xóa kết quả kiểm tra của phiên.

**Quản trị viên** (tab **⚙️ Quản trị**)
* Thêm, sửa, xóa tài khoản; đặt lại mật khẩu; **Hủy ghi nhớ** để buộc một tài khoản đăng nhập lại trên mọi máy.
* Xem và tải nhật ký sử dụng: mỗi lần kiểm tra ghi tên tệp Word, số bảng, số sai lệch, số cần xem xét và
  thời gian xử lý; không ghi nội dung báo cáo.
"""

TROUBLESHOOTING = """
* **Không tải được tệp:** kiểm tra đúng đuôi .docx và tệp không bị khóa hoặc đang mở ở chế độ chỉ đọc.
* **Thông báo không xử lý được tệp:** mở lại tệp bằng Word, lưu lại dạng .docx rồi thử lại.
* **Đã có Excel nhưng không có Word comment:** vẫn dùng được kết quả Excel; thử lại với bản .docx đã lưu lại.
* **Xử lý lâu:** báo cáo lớn cần vài chục giây; không tải lại trang khi đang xử lý.
* **Tải lại trang (F5)** làm mất kết quả kiểm tra của phiên; nếu không bật ghi nhớ thì phải đăng nhập lại.
* **Bản cloud ngủ đông:** lần truy cập đầu sau thời gian dài cần đợi 1–2 phút.
* **Quên mật khẩu:** liên hệ quản trị viên để được đặt lại.
"""

_SECTIONS = (
    ("🚀 Bắt đầu nhanh", QUICK_START),
    ("📂 Kiểm tra báo cáo", PROCESS),
    ("📊 Đọc kết quả Excel", READ_EXCEL),
    ("💬 Rà soát Word", READ_WORD),
    ("👤 Tài khoản", ACCOUNT),
    ("🛠️ Sự cố thường gặp", TROUBLESHOOTING),
)


def guide_sections() -> tuple[tuple[str, str], ...]:
    return _SECTIONS


def show_user_guide() -> None:
    st.dialog("📘 Hướng dẫn sử dụng", width="large")(_render_guide)()


def _render_guide() -> None:
    tabs = st.tabs([title for title, _ in _SECTIONS])
    for tab, (_, content) in zip(tabs, _SECTIONS):
        with tab:
            st.markdown(content)
