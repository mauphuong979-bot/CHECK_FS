# AGENTS.md

## Phạm vi và tính bắt buộc

File này là bộ quy chuẩn bắt buộc đối với Codex và mọi AI coding agent khi làm việc trong repository này.

Các quy định áp dụng cho mọi hoạt động tạo mới, cập nhật, sửa lỗi, refactor, tối ưu hoặc bổ sung tính năng. Trước khi chỉnh sửa code, agent phải đọc:

1. `AGENTS.md`.
2. `PROJECT_STRUCTURE.md`.
3. Entry point, module chịu trách nhiệm, dependency và test liên quan đến phạm vi thay đổi.

Nếu các nguyên tắc xung đột, thứ tự ưu tiên là:

1. Đúng dữ liệu.
2. Đúng logic nghiệp vụ kiểm toán.
3. Bảo mật và không làm mất dữ liệu.
4. Không phá vỡ hành vi hiện tại ngoài phạm vi yêu cầu.
5. Hiệu năng xử lý.
6. Tiết kiệm bộ nhớ.
7. Dễ bảo trì và khả năng mở rộng.
8. Chất lượng giao diện và trải nghiệm người dùng.

Không được đánh đổi tính chính xác của dữ liệu hoặc logic nghiệp vụ để lấy tốc độ, sự thuận tiện khi lập trình hoặc hình thức giao diện.

## Mục tiêu dự án

Dự án hỗ trợ kiểm toán viên kiểm tra tính chính xác về mặt số học và tính nhất quán của số liệu trên Báo cáo tài chính.

Các chức năng chính:

- Kiểm tra tổng cộng dọc trên Báo cáo tài chính.
- Kiểm tra tổng cộng ngang trên Báo cáo tài chính.
- Kiểm tra tính logic giữa Bảng cân đối kế toán (BS), Báo cáo kết quả hoạt động kinh doanh (PL), Báo cáo lưu chuyển tiền tệ (CF) và Thuyết minh Báo cáo tài chính.
- Đối chiếu chéo số liệu giữa các biểu mẫu.
- Phát hiện sai lệch số học.
- Xuất kết quả kiểm tra ra báo cáo Excel.

Người sử dụng chính: kiểm toán viên.

## 1. Thiết kế như một ứng dụng chuyên nghiệp

- Giao diện phải có bố cục rõ ràng, cân đối, hiện đại và nhất quán.
- Thiết kế theo tư duy của một sản phẩm thực tế; không tạo cảm giác demo, prototype hoặc code tạm.
- Phải có phân cấp thị giác rõ ràng giữa tiêu đề, nội dung chính, thông tin phụ và hành động của người dùng.
- Khoảng cách, padding, margin, kích thước chữ, button, form, table, card và các thành phần UI phải đồng bộ.
- Ưu tiên giao diện sạch, dễ đọc, dễ thao tác và chuyên nghiệp đối với kiểm toán viên.
- Giao diện phải responsive tốt trên các kích thước màn hình phổ biến trong giới hạn của Streamlit.
- Khi bổ sung tính năng mới, phải chủ động bố trí lại giao diện trong phạm vi hợp lý nếu cần để tính năng hòa nhập tự nhiên với tổng thể ứng dụng.
- Không được coi phần hiển thị là hoàn tất nếu còn lỗi căn chỉnh, tràn nội dung, phân cấp thị giác không rõ hoặc thao tác khó hiểu.
- Toàn bộ nội dung hiển thị cho người dùng phải sử dụng tiếng Việt, trừ tên kỹ thuật hoặc thuật ngữ cần giữ nguyên.

## 2. Duy trì tính nhất quán

- Trước khi tạo component hoặc style mới, phải kiểm tra component/style hiện có để tái sử dụng.
- Không tạo nhiều cách thiết kế khác nhau cho cùng một loại chức năng.
- Giữ nhất quán về màu sắc, typography, icon, button, form control, modal, table, notification và navigation.
- Tính năng mới phải tuân theo design language hiện có của ứng dụng.
- Nếu chưa có design system chính thức, phải suy ra pattern từ giao diện hiện tại và áp dụng nhất quán; chỉ tạo pattern mới khi thật sự cần thiết.
- Các thông báo cùng cấp độ như info, warning, error và success phải có cách trình bày và ngôn ngữ nhất quán.

## 3. Cấu trúc code chuyên nghiệp

- Code phải rõ ràng, dễ đọc, dễ bảo trì và có cấu trúc hợp lý.
- Tôn trọng kiến trúc hiện tại: tách biệt UI, service điều phối, I/O, extraction, normalization, domain và reporting.
- Tách component, function, service, utility hoặc module khi cần thiết; tránh tạo file, class hoặc function có quá nhiều trách nhiệm.
- Không tạo duplicate code. Tái sử dụng abstraction hiện có khi phù hợp.
- Đặt tên biến, function, class, component và file rõ nghĩa, nhất quán với convention hiện tại.
- Hạn chế hard-code; đưa giá trị dùng chung thành constant hoặc configuration khi điều đó làm code rõ và ổn định hơn.
- Không thực hiện thay đổi không liên quan đến yêu cầu hiện tại.
- Không over-engineer; ưu tiên giải pháp đơn giản, rõ ràng, ổn định nhưng có khả năng mở rộng hợp lý.
- Không phá vỡ public interface nếu không thật sự cần thiết và chưa đánh giá đầy đủ caller.
- Không tạo dependency vòng hoặc đưa logic nghiệp vụ vào layer UI.
- Bổ sung type hint cho code Python mới hoặc được sửa khi phù hợp với convention hiện tại.
- Chỉ bổ sung docstring/comment khi chúng giải thích contract, quyết định hoặc logic không hiển nhiên; không viết comment lặp lại code.
- Không giữ code chết, import không dùng hoặc nhánh tạm sau khi hoàn tất thay đổi.

## 4. UX đầy đủ

Đối với mọi chức năng có tương tác với người dùng, phải xem xét và triển khai các trạng thái phù hợp:

- Loading.
- Empty state.
- Error state.
- Success state.
- Disabled state.
- Validation.
- Confirmation đối với thao tác quan trọng, không thể hoàn tác hoặc có khả năng gây mất dữ liệu.

Yêu cầu bổ sung:

- Thông báo lỗi phải dễ hiểu với người dùng, nêu được điều họ có thể làm tiếp theo khi phù hợp; không chỉ hiển thị lỗi kỹ thuật hoặc stack trace.
- Không để người dùng kích hoạt lặp lại một tác vụ đang xử lý nếu việc đó có thể tạo xung đột hoặc tiêu tốn tài nguyên không cần thiết.
- Empty state phải hướng dẫn hành động tiếp theo thay vì chỉ để trống.
- Validation phải diễn ra càng gần điểm nhập liệu càng tốt và không tự ý sửa số liệu kế toán của người dùng.
- Success state phải xác nhận rõ kết quả đã tạo và cung cấp hành động tiếp theo phù hợp.

## 5. Không làm giảm chất lượng phần hiện có

Khi chỉnh sửa một chức năng:

- Kiểm tra ảnh hưởng đến các chức năng, module, interface, rule và test liên quan.
- Không phá vỡ hành vi hiện tại nếu yêu cầu không đề nghị thay đổi.
- Không làm giao diện kém nhất quán hơn sau khi thêm tính năng.
- Nếu code hoặc bố cục liên quan trực tiếp gây khó khăn cho việc triển khai, có thể refactor trong phạm vi hợp lý để giữ kiến trúc sạch.
- Không mở rộng refactor sang khu vực không liên quan chỉ vì phát hiện cơ hội cải thiện.
- Giữ backward compatibility khi có thể; nếu không thể, phải giải thích tác động trước khi thực hiện thay đổi lớn.
- Mọi thay đổi liên quan công thức kiểm toán phải được đối chiếu với dữ liệu gốc và rule đã được xác nhận.

## 6. Kiểm tra sau khi thay đổi

Sau mỗi lần sửa code, agent phải:

- Kiểm tra syntax và type nếu project hỗ trợ.
- Chạy lint, test và build phù hợp nếu repository có công cụ tương ứng.
- Kiểm tra các luồng chính bị ảnh hưởng, bao gồm trạng thái thành công và lỗi.
- Sửa các lỗi phát sinh trực tiếp từ thay đổi vừa thực hiện.
- Kiểm tra import và entry point khi thay đổi cấu trúc module.
- Với thay đổi tạo Excel, xác minh workbook sinh ra, sheet, formula, style, liên kết và dữ liệu đầu ra trong phạm vi ảnh hưởng.
- Với thay đổi UI, kiểm tra bố cục và trải nghiệm ở kích thước màn hình phù hợp; không chỉ xác nhận rằng code chạy.
- Không coi công việc hoàn tất nếu giao diện hoạt động nhưng bố cục còn lỗi rõ ràng hoặc trải nghiệm chưa đầy đủ.
- Báo rõ test/lint/build nào đã chạy, kết quả và kiểm tra nào chưa chạy được cùng nguyên nhân.
- Cập nhật `PROJECT_STRUCTURE.md` nếu thay đổi cấu trúc, entry point, interface, dependency, configuration hoặc luồng dữ liệu.

Không được tuyên bố test pass nếu công cụ không được cài, test không được chạy hoặc kết quả chưa được xác minh.

## 7. Nguyên tắc mặc định khi yêu cầu chưa mô tả UI chi tiết

- Khi yêu cầu chỉ mô tả chức năng, Codex phải chủ động chọn cách trình bày hợp lý nhất dựa trên giao diện, design language và kiến trúc hiện tại.
- Không cần hỏi lại người dùng về các quyết định UI nhỏ mà một lập trình viên có kinh nghiệm có thể tự quyết định hợp lý.
- Chỉ yêu cầu xác nhận khi quyết định có thể thay đổi đáng kể workflow, nghiệp vụ, dữ liệu, quyền truy cập, hành vi hiện tại hoặc phạm vi yêu cầu.
- Ưu tiên phương án ít gây bất ngờ, dễ khám phá, dễ hoàn tác và phù hợp với người dùng là kiểm toán viên.

Mục tiêu cuối cùng: **mọi phần code được thêm hoặc chỉnh sửa phải tạo cảm giác đây là một ứng dụng được xây dựng và duy trì bởi một đội ngũ lập trình chuyên nghiệp, thay vì tập hợp các tính năng được bổ sung rời rạc qua từng yêu cầu.**

## Môi trường triển khai

Ứng dụng được thiết kế để triển khai trên Streamlit Cloud.

Khi phát triển:

- Hỗ trợ nhiều người dùng đồng thời.
- Không sử dụng giải pháp chỉ hoạt động trên Windows.
- Hạn chế phụ thuộc vào đường dẫn cục bộ.
- Ưu tiên thư mục tạm và API đa nền tảng.
- Bảo đảm hoạt động trên Linux của Streamlit Cloud.
- Tối ưu thời gian xử lý, bộ nhớ và tài nguyên hệ thống.
- Không dùng global mutable state để lưu dữ liệu riêng của người dùng.
- Không giả định filesystem runtime là nơi lưu trữ lâu dài.

## Bảo mật và dữ liệu khách hàng

- Không lưu dữ liệu khách hàng lâu dài trên máy chủ.
- Phải xóa tệp tạm ở cả luồng thành công và luồng lỗi.
- Không ghi dữ liệu nhạy cảm, nội dung báo cáo hoặc thông tin nhận dạng khách hàng vào log.
- Không hiển thị thông tin hệ thống, đường dẫn nội bộ hoặc stack trace cho người dùng cuối.
- Không hard-code mật khẩu, token, API key hoặc secret trong source code.
- Sử dụng Streamlit Secrets hoặc biến môi trường để quản lý bí mật nếu phát sinh nhu cầu.
- Hạn chế tối đa việc lộ source code hoặc cấu hình nội bộ.
- Không đưa dữ liệu thật vào fixture/test mới nếu chưa được ẩn danh phù hợp.

## Quy định về Python

- Ưu tiên DuckDB khi xử lý dữ liệu lớn.
- Chỉ sử dụng Pandas khi kích thước dữ liệu phù hợp.
- Tránh nạp toàn bộ dữ liệu vào RAM nếu không cần thiết.
- Ưu tiên xử lý theo từng phần (chunk) khi định dạng và luồng nghiệp vụ cho phép.
- Thiết kế module hóa và giữ ranh giới trách nhiệm rõ ràng.
- Tách biệt giao diện, xử lý dữ liệu và nghiệp vụ.
- Ưu tiên code dễ đọc, dễ kiểm thử và dễ bảo trì.
- Khi xử lý số kế toán, ưu tiên kiểu dữ liệu giữ được độ chính xác; phải đánh giá trước khi chuyển `Decimal` sang `float`.
- Bảo toàn Unicode tiếng Việt và kiểm tra các trường hợp dấu phân cách số.
- Không dùng bare `except`; exception catch-all tại biên UI phải trả thông báo an toàn và có cơ chế chẩn đoán phù hợp không lộ dữ liệu nhạy cảm.

## Quy định về VBA

Nếu công việc có chỉnh sửa hoặc tạo VBA:

- Bắt buộc sử dụng `Option Explicit`.
- Đặt tên biến có ý nghĩa.
- Hạn chế `Select` và `Activate`.
- Ưu tiên xử lý bằng mảng thay vì thao tác trực tiếp trên Worksheet.
- Hạn chế đọc/ghi dữ liệu từng ô.
- Luôn cân nhắc hiệu năng với dữ liệu lớn.
- Nếu sử dụng tiếng Việt trong VBA, ưu tiên `ChrW()` để tránh lỗi Unicode trên các môi trường khác nhau.
- Không giữ đoạn mã dư thừa hoặc không còn sử dụng.

## Quy định về dữ liệu đầu vào

Các định dạng mục tiêu của dự án:

- XLSX.
- XLSB.
- CSV.
- TXT.
- DOCX.
- DOC.
- DOCM.

Agent phải phân biệt rõ giữa **định dạng mục tiêu** và **định dạng source hiện đã triển khai**. Không được tuyên bố một định dạng đã được hỗ trợ chỉ vì extension xuất hiện trong cấu hình. Mỗi định dạng mới phải có reader, validation, xử lý lỗi và test phù hợp.

Ứng dụng cần tự động nhận diện định dạng phù hợp khi chức năng đó được triển khai.

## Quy định về dữ liệu đầu ra

Định dạng đầu ra ưu tiên của dự án là XLSB. Ngoài ra có thể hỗ trợ XLSX và CSV.

Agent phải mô tả đúng định dạng source hiện có thể tạo; không được đổi extension mà không có writer thực tế tương ứng.

Kết quả kiểm tra phải:

- Rõ ràng, dễ theo dõi và phục vụ công tác kiểm toán.
- Cho phép truy vết từ sai lệch về bảng, dòng, cột hoặc mã chỉ tiêu liên quan.
- Không tự ý hiệu chỉnh dữ liệu kế toán gốc.
- Phân biệt rõ lỗi xác định được và trường hợp cần kiểm toán viên xem xét.

## Quy định chuyển đổi dữ liệu số

- Tự động nhận diện chuỗi số khi có đủ cơ sở.
- Chuyển chuỗi số thành giá trị số khi phù hợp.
- Nhận diện định dạng số theo dữ liệu/thiết lập Excel trong giới hạn đã triển khai.
- Hỗ trợ dấu phân cách hàng nghìn, số âm và định dạng kế toán.
- Hạn chế tối đa chuyển đổi sai kiểu dữ liệu.
- Không mặc định giá trị không parse được là số hợp lệ nếu điều đó có thể che giấu lỗi dữ liệu.
- Mọi thay đổi heuristic parse số phải có test cho trường hợp Việt Nam, quốc tế, số âm, số thập phân, giá trị trống và văn bản không phải số.

## Quy định ghi log

Mỗi lần xử lý cần ghi nhận khi cơ chế logging được triển khai:

- Thời gian bắt đầu.
- Thời gian kết thúc.
- Thời gian thực hiện.
- Số lượng tệp xử lý.
- Các lỗi phát sinh ở mức đủ để chẩn đoán.

Không ghi dữ liệu nhạy cảm của khách hàng vào log. Log cho người vận hành và thông báo cho người dùng phải được tách biệt về mức độ chi tiết.

## Quy tắc nghiệp vụ kiểm toán

- Mọi phép kiểm tra số học phải ưu tiên dựa trên dữ liệu gốc từ báo cáo tài chính.
- Không tự ý thay đổi quy tắc kiểm tra, hệ số, dấu, mã chỉ tiêu hoặc công thức nghiệp vụ nếu chưa được người dùng xác nhận.
- Nếu có nhiều cách hiểu về một chỉ tiêu, phải yêu cầu người dùng xác nhận trước khi tự động kết luận.
- Không tự ý suy diễn, làm tròn, bù trừ hoặc hiệu chỉnh số liệu kế toán.
- Phải giữ nhất quán giữa phép tính trong Python và công thức ghi vào Excel.
- Khi thiếu dữ liệu hoặc mã chỉ tiêu, phải thể hiện rõ là thiếu dữ liệu/cần xác minh thay vì âm thầm coi là đúng.

## Khi rà soát hoặc sửa code

Luôn đánh giá:

- Độ chính xác dữ liệu và logic nghiệp vụ.
- Hiệu năng và bộ nhớ.
- Xử lý lỗi và đầy đủ trạng thái UX.
- Unicode và định dạng số.
- Khả năng mở rộng và bảo trì.
- Khả năng hoạt động trên Streamlit Cloud/Linux.
- Ảnh hưởng đến module, test, public interface và dữ liệu đầu ra liên quan.

Trước thay đổi lớn, phải giải thích ngắn gọn:

- Mục đích thay đổi.
- Lợi ích.
- Rủi ro.
- Ảnh hưởng đến chức năng hiện tại.
- Phương án kiểm tra sau thay đổi.

## Phong cách trao đổi

- Sử dụng tiếng Việt.
- Trình bày ngắn gọn, rõ ràng và dựa trên bằng chứng từ source.
- Giải thích nguyên nhân khi đề xuất thay đổi.
- Nêu tác động đến hiệu năng, bộ nhớ, nghiệp vụ hoặc UX nếu có.
- Không thay đổi phần không liên quan.
- Ưu tiên giải pháp đơn giản, ổn định và dễ bảo trì.
- Ghi rõ nội dung `Chưa xác định được` hoặc `Cần xác minh` khi source chưa đủ căn cứ.
