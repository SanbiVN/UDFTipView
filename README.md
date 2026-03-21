# UDFTipView
⭐HÀM THIẾT LẬP HÀNH ĐỘNG CHUỘT HIỂN THỊ CỬA SỔ TIP CHO EXCEL

với Hàm **TipView**

![TipView2](https://github.com/user-attachments/assets/fb902b17-5dfb-4c03-838a-17ae8e13ca7e)


### ✳Ưu điểm và chức năng

1. Chỉ cần gõ hàm để thiết lập nhanh chóng.​
2. Hiển thị Tip khi rê chuột chọn ô hoặc chỉ rê chuột thời gian thực.​
3. Hiển thị Tip gồm: Chú thích, vùng ô (Range), Ảnh (trang tính, tệp gốc, tệp thư mục) và Biểu đồ.​
4. Không làm mất trạng thái Undo, Redo và cả Clipboard (Bộ nhớ tạm).​

### ✳Hướng dẫn sử dụng hàm TipView

Hàm: ```=TipView(Vùng_sự_kiện_chuột, Các_đối_số_thiết_lập...)```

Cách viết hàm nhanh, gõ vào ô chuỗi =TipView và ấn tổ hợp phím Ctrl+Shift+A

Tham số :
​
❄Vị trí |	Tham số	Kiểu |	Diễn giải
-------- | ---------- |---------------------------
1 |	Vùng_sự_kiện_chuột	 | Vùng ô hoặc Name, Vùng khi chuột rê vào sẽ hiển thị Tip
2 |	Các_đối_số_thiết_lập	 |Các hàm đối số bổ trợ,	Có thể nhập nhiều đối số phía sau, để thiết lặp

Gõ hàm TipView_HuongDan() để hiển thị hướng dẫn đầy đủ khi cần.

# ✳Các hàm dưới đây thiết lập để hiển thị cửa sổ Tip, và chúng phải được gõ trong hàm TipView
​
❄Các hàm thiết lập |	Kiểu
-------- | ----------
**TipView_Data**(Vùng_dữ_liệu) |	Nhập vùng dữ liệu để TipView_GetRow, và TipView_GetColumn tham chiếu
**TipView_ViewRange**(Vùng_ô, ...) |	Vùng ô, có thể nhập nhiều hơn, tip hiển thị tối đa 3 vùng ô
**TipView_ViewImage**(Tên/Path, Chiều_rộng, chiều_cao) |	Hiển thị ảnh gồm: Ảnh PNG, GIF, JPG, JPGE, BITMAP, ICO<br> Nguồn tham chiếu gồm: <br>- Ảnh trong trang tính: **TipView_ViewImage**("Name...1")<br> - Ảnh nằm trong đóng gói tệp (xl/Embeding, xl/CustomUI/Images): Nhập như trên, tự động tìm kiếm.<br>- Ảnh từ đường dẫn thư mục: **TipView_ViewImage**("C:\Folder\Image...1.png")<br> - Ảnh từ liên kết URL: **TipView_ViewImage**("https:\\...Image...1.png")
**TipView_ViewCharts**(Biểu_đồ, ...) |	Biểu đồ, có thể nhập nhiều hơn, tip hiển thị tối đa 3 biểu đồ
**TipView_ViewNote**(Ghi_chú) |	Ghi chú sẽ hiển thị
**TipView_MouseOver**() |	Tùy chọn rê chuột thì hiển thị tip thời gian thực thay cho chọn ô
**TipView_ShowVertical**() |	Tùy chọn hiển thị theo chiều dọc, nếu có 2 cửa sổ tip trở lên, mặc định là chiều ngang.


🌟Để có ngay biểu thức nhanh, đầy đủ ví dụ, hãy nhập chỉ tên hàm:​
> =TipView_ViewRange() ​
> hoặc =TipView_ViewImage() ​
> hoặc =TipView_ViewCharts()

# ✳Các hàm trả về vị trí và thứ tự dòng, cột hoặc ô hiện tại để tham chiếu dữ liệu
Các hàm này cần tham chiếu cho dữ liệu sẽ hiển thị trong Tip gồm ViewRange và ViewCharts​
Nếu nhập TipView_ViewRange(H1:J5), thì trong vùng H1:J5 nhập biểu thức phải chứa các hàm bên dưới để tham chiếu\
✨Ví dụ:
​
```=XLOOKUP(TipView_GetRow(2,TRUE),$B$14:$B$25,$A$14:$A$25,0,0,1)```​
​
❄Các hàm |	Kiểu
-------- | ----------
**TipView_GetRow**(Cột_tham_chiếu,[Trả_về_ô_đối_tượng]) |	Trả về thứ tự dòng hoặc ô hiện tại đang rê chuột trong bảng, phụ thuộc Vùng_dữ_liệu nhập trong TipView
**TipView_GetColumn**(Dòng_tham_chiếu,[Trả_về_ô_đối_tượng]) |	Trả về thứ tự cột hoặc ô hiện tại đang rê chuột trong bảng, phụ thuộc Vùng_dữ_liệu nhập trong TipView

​
✨Ví dụ với thiết lập:​
```=TipView(A2:A10,TipView_Data(A1:E10),TipView_ViewRange(H1:J5))```
​
Khi gọi: TipView_GetRow(2,TRUE) thì sẽ trả về ô đối tượng dòng trong cột 2 của vùng ô dữ liệu A1:E10

# ✳Các hàm thiết lập vị trí hiển thị cửa sổ Tip quanh ô chọn​
> Cửa sổ tip mặc định vị trí hiển thị ở phía dưới bên phải ô hiện hành. Các hàm sau đây sẽ giúp thay đổi vị trí linh hoạt hơn.​

✡ Vị trí bắt đầu nằm quanh ô:​
​
❄Thông số vị trí |	Diễn giải
--------------- | -------------------------------------------
**Tip_LeftTop**  |	Vị trí bên trái + phía trên
**Tip_LeftBottom**  |	Vị trí bên trái + phía dưới
**Tip_RightTop** |	Vị trí bên phải + phía trên
**Tip_RightBottom**  |	Vị trí bên phải + phía dưới

✡ Vị trí cửa sổ sẽ hiển thị:​
​
❄Thông số vị trí	 | Diễn giải
--------------- | -------------------------------------------
**Tip_WindowRightBelow**  |	Cửa sổ sẽ nằm ở bên phải + phía dưới
**Tip_WindowRightAbove** |	Cửa sổ sẽ nằm ở bên phải + phía trên
**Tip_WindowLeftBelow** |	Cửa sổ sẽ nằm ở bên trái + phía dưới
**Tip_WindowLeftAbove** |	Cửa sổ sẽ nằm ở bên trên + phía trên

#### ✨Ví dụ với thiết lập:
​
```=TipView(A2:A10,TipView_ViewRange(H1:J5), Tip_LeftTop(),Tip_WindowRightAbove())``` \
​
Thì cửa sổ sẽ bắt đầu bên trái + phía trên của ô chọn, cửa sổ sẽ nằm ở bên phải + phía trên bắt đầu từ vị trí đó\​

(Mã VBA sẽ sớm cập nhật phiên bản đầu tiên)
