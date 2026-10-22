# Tracking ngày – Team SS Pháp (GT, NPP Tâm Hạnh Đà Nẵng)

Dashboard chỉ tiêu ngày theo DDKD. Trang xem: `index.html` (GitHub Pages).

## Quy trình nạp số liệu
- **Sáng / trưa / chiều (tracking trong ngày):** nạp file đơn mới (BÁO CÁO LINE ITEM) để xem thực hiện. Không lưu vào `data/`.
- **Tối (chính thức):** nạp file đơn của ngày và lưu vào `data/orders_YYYY-MM-DD.csv.gz`.
- File sau luôn thay toàn bộ các ngày mà nó phủ, nên nạp lại trong ngày không bị cộng đôi. Đơn đã hủy bị loại.

## Lệnh dựng lại
```
# tracking trong ngày
python3 src/build.py --src inputs --data data --orders <file_don.xlsx> --r3 inputs/CT_TH_789.xlsx --sched inputs/Visit_Schedule.xlsx --standalone --out index.html
# nạp chính thức buổi tối (thêm --save)
python3 src/build.py --src inputs --data data --orders <file_don.xlsx> --save --r3 inputs/CT_TH_789.xlsx --sched inputs/Visit_Schedule.xlsx --standalone --out index.html
```
