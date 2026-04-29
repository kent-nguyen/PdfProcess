# XuLyPdf — Chuyển đổi Sao kê PDF sang Excel

Công cụ tự động chuyển sao kê ngân hàng dạng PDF (hiện hỗ trợ BIDV) thành file Excel có cấu trúc, sử dụng OCR để đọc từng ô trong bảng.

---

## Lệnh hay dùng

```bash
# Chạy toàn bộ pipeline cho một dải trang
python Run.py --bank bidv --pages 161-180

# Kết hợp tất cả các trang vào một file Output.xlsx
python CombineExcels.py

# Chạy từng bước thủ công cho một trang
python ConvertToImages.py --pages 105
# Chỉnh độ nghiêng ảnh (nếu cần)
python Deskew.py --page 105
# Cắt hình ra thành các ô
python SplitTableCells.py --page 105
# Chuyển ô thành file raw_105.png
python CellsToExcel.py --bank bidv --page 105
# Đọc file raw và xuất ra file 105.png
python RefineData.py --bank bidv --page 103

# Tách dòng kép thủ công (chỉ định thư mục Cells, dòng bị kép, và vị trí cắt Y)
python SplitDoubledRow.py --cells-dir .\Pages\136\Cells\ --row 5 --split-y 162
```

---

## Yêu cầu

- Python 3.9+
- [Poppler](https://github.com/oschwartz10612/poppler-windows/releases/) (cần có trong PATH, hoặc khai báo đường dẫn trong `ConvertToImages.py`)
- Các thư viện Python: xem `pyproject.toml` (quản lý bằng Poetry)

Cài đặt dependencies:

```bash
poetry install
```

---

## Quy trình làm việc

Có thể chạy **toàn bộ bước 1–5 một lần** bằng `Run.py`.

```
Source.pdf
    │
    ▼ (1) ConvertToImages.py
Pages/<trang>/<trang>.png
    │
    ▼ (2) Deskew.py
Pages/<trang>/<trang>.png  (ghi đè nếu ảnh bị nghiêng)
    │
    ▼ (3) SplitTableCells.py
Pages/<trang>/Cells/row000_col000.png  …
    │
    │   [Tuỳ chọn] SplitDoubledRow.py  ← sửa thủ công nếu có dòng kép
    │
    ▼ (4) CellsToExcel.py
Pages/<trang>/raw_<trang>.xlsx
    │
    ▼ (5) RefineData.py
Pages/<trang>/<trang>.xlsx
    │
    ▼ (6) CombineExcels.py
Output.xlsx
```

---

## Mô tả các script

### `Run.py` — Chạy toàn bộ pipeline (bước 1–4)

Chạy lần lượt 4 bước chính cho một hoặc nhiều trang, sau đó in tóm tắt lỗi.

```bash
python Run.py --pages 1-5 --bank bidv
python Run.py --pages 1,3,7-10 --bank bidv
```

| Tham số   | Mặc định     | Mô tả                                               |
| --------- | ------------ | --------------------------------------------------- |
| `--pages` | _(bắt buộc)_ | Trang cần xử lý: `1`, `1,3,5`, `2-5`, `1-3,7,10-12` |
| `--bank`  | `bidv`       | Mã ngân hàng, chọn cấu hình từ `banks/<bank>.py`    |

---

### `ConvertToImages.py` — Bước 1: PDF → PNG

Chuyển từng trang PDF thành ảnh PNG độ phân giải cao (300 DPI), tự động tăng tương phản và làm nét để OCR chính xác hơn.

Đầu ra: `Pages/<trang>/<trang>.png`

```bash
python ConvertToImages.py --pages 1-5
python ConvertToImages.py            # chuyển tất cả các trang
```

> **Lưu ý:** File PDF gốc phải đặt tên là `Source.pdf` và để ở thư mục gốc của project.

---

### `Deskew.py` — Bước 2: Chỉnh độ nghiêng ảnh

Phát hiện và chỉnh độ nghiêng của ảnh trang bằng cách tìm đường thẳng ngang qua Hough Transform, rồi xoay ảnh về đúng góc. Ghi đè trực tiếp lên ảnh gốc nếu phát hiện nghiêng.

Đầu ra: `Pages/<trang>/<trang>.png` (ghi đè nếu ảnh bị nghiêng)

```bash
python Deskew.py --page 3
python Deskew.py               # xử lý tất cả các trang hiện có
python Deskew.py --page 3 --debug  # lưu thêm ảnh binary và edges để kiểm tra
```

---

### `SplitTableCells.py` — Bước 3: Phát hiện bảng và cắt ô

Phân tích ảnh PNG để tìm các đường kẻ ngang/dọc của bảng, sau đó cắt từng ô thành ảnh riêng biệt.

Đầu ra: `Pages/<trang>/Cells/row000_col000.png`, `row000_col001.png`, …

```bash
python SplitTableCells.py --page 3
python SplitTableCells.py            # xử lý tất cả các trang
python SplitTableCells.py --page 3 --debug  # lưu thêm ảnh debug đường kẻ
```

---

### `SplitDoubledRow.py` — Công cụ sửa thủ công: Tách dòng kép

Dùng khi một ô chứa **hai dòng dữ liệu bị gộp chung** (hai số STT trong một ô). Script tự động phát hiện đường kẻ nội bộ, tách ô ra và đánh số lại toàn bộ các dòng phía sau.

```bash
python SplitDoubledRow.py --cells-dir Pages/97/Cells --row 19
```

| Tham số       | Mô tả                                           |
| ------------- | ----------------------------------------------- |
| `--cells-dir` | Đường dẫn đến thư mục `Cells` của trang cần sửa |
| `--row`       | Chỉ số dòng bị kép (đếm từ 0)                   |

---

### `CellsToExcel.py` — Bước 4: OCR ô → Excel thô

Dùng EasyOCR để đọc chữ trong từng ảnh ô, ghi kết quả vào file Excel thô.

Đầu ra: `Pages/<trang>/raw_<trang>.xlsx`

```bash
python CellsToExcel.py --page 3 --bank bidv
python CellsToExcel.py               # xử lý tất cả các trang
```

Mỗi cột được OCR với bộ ký tự cho phép riêng (ví dụ: cột số tiền chỉ nhận chữ số và dấu phẩy/chấm), giúp giảm lỗi OCR.

---

### `RefineData.py` — Bước 5: Chuẩn hoá và sửa lỗi dữ liệu

Đọc file Excel thô, áp dụng các bộ định dạng và sửa lỗi theo cấu hình ngân hàng:

- **Formatters:** chuẩn hoá định dạng ngày giờ, số tiền, số thứ tự, mô tả giao dịch, tên ngân hàng đối ứng.
- **Fixers:** tự động sửa số thứ tự bị lỗi OCR, kiểm tra tính đúng đắn của số dư (Debit/Credit/Balance).
- Thêm 3 cột phụ vào cuối: `Fixed`, `Error`, `Notes` để đánh dấu các dòng đã được sửa hoặc còn lỗi.

Đầu ra: `Pages/<trang>/<trang>.xlsx`

```bash
python RefineData.py --page 3 --bank bidv
python RefineData.py                 # xử lý tất cả các trang
```

---

### `CombineExcels.py` — Bước 5: Gộp tất cả trang thành một file

Gộp tất cả file `Pages/<trang>/<trang>.xlsx` thành một file `Output.xlsx` duy nhất. Tự động thêm cột `Page` ở đầu để biết dữ liệu thuộc trang nào.

```bash
python CombineExcels.py
python CombineExcels.py --input Pages --output Output.xlsx
```

---

## Xử lý lỗi thường gặp

### Hình bị nghiêng

Nếu ảnh trang bị nghiêng, `SplitTableCells.py` sẽ không phát hiện đúng đường kẻ bảng và cắt ô sai. Chạy `Deskew.py` để chỉnh tự động:

```bash
python Deskew.py --page 95
```

Thêm `--debug` để lưu ảnh binary và edges ra kiểm tra nếu kết quả chưa đúng. Sau khi ảnh đã thẳng, chạy lại từ bước 3 cho trang đó:

```bash
python SplitTableCells.py --page 95
python CellsToExcel.py --page 95 --bank bidv
python RefineData.py --page 95
```

---

### Dòng bị gộp đôi (doubled row)

Xảy ra khi hai giao dịch liền kề bị gom vào một dòng duy nhất trong ảnh. Triệu chứng: hai số STT xuất hiện trong cùng một ô sau bước 2.

Chạy `SplitDoubledRow.py` để tách, chỉ định đúng thư mục `Cells` và chỉ số dòng bị kép (đếm từ 0):

```bash
python SplitDoubledRow.py --cells-dir Pages/97/Cells --row 19
```

Sau đó chạy lại từ bước 3 cho trang đó:

```bash
python CellsToExcel.py --page 97 --bank bidv
python RefineData.py --page 97
```

> **Lưu ý:** Chỉ chạy `SplitDoubledRow.py` **một lần** cho mỗi dòng bị kép. Chạy lại sẽ shift thêm một lần nữa, tạo ra dòng trùng lặp ở cuối bảng.

---

## Cấu hình ngân hàng (`banks/`)

Mỗi file trong thư mục `banks/` định nghĩa cấu hình cho một ngân hàng:

| Biến                | Mô tả                                                     |
| ------------------- | --------------------------------------------------------- |
| `COLUMN_ALLOWLISTS` | Bộ ký tự cho phép khi OCR từng cột (giảm lỗi nhận dạng)   |
| `FORMATTERS`        | Danh sách hàm chuẩn hoá dữ liệu từng cột                  |
| `FIXERS`            | Danh sách hàm tự động sửa lỗi logic (STT, số dư, …)       |
| `GARBAGE_DATE_COLS` | Cột ngày dùng để phát hiện và bỏ các dòng rác ở cuối bảng |

Hiện tại hỗ trợ: **`bidv`**

---

## Cấu trúc thư mục sau khi chạy

```
XuLyPdf/
├── Source.pdf              ← File PDF đầu vào (đặt tay)
├── Pages/
│   ├── 1/
│   │   ├── 1.png           ← Ảnh trang (bước 1)
│   │   ├── Cells/          ← Ảnh từng ô (bước 2)
│   │   ├── raw_1.xlsx      ← Excel thô từ OCR (bước 3)
│   │   └── 1.xlsx          ← Excel đã chuẩn hoá (bước 4)
│   └── 2/ …
└── Output.xlsx             ← File tổng hợp cuối cùng (bước 5)
```

---

## Hướng dẫn cài đặt

### Poppler (Windows)

1. Tải bản release mới nhất tại: https://github.com/oschwartz10612/poppler-windows/releases/
2. Giải nén, ví dụ vào `C:\poppler`
3. Thêm đường dẫn `C:\poppler\Library\bin` vào biến môi trường **PATH**:
   - Mở **Start** → tìm **"Edit the system environment variables"**
   - Bấm **Environment Variables…** → chọn **Path** → **Edit** → **New**
   - Dán vào `C:\poppler\Library\bin` → OK
4. Mở lại terminal và kiểm tra: `pdftoppm -v`

### Poetry

```bash
pip install poetry
```

Kiểm tra:

```bash
poetry --version
```

### Cài dependencies của project

Trong thư mục project, chạy:

```bash
poetry install
```

> **Lưu ý:** Lần chạy đầu tiên EasyOCR sẽ tự tải model tiếng Anh (~100 MB), máy cần có kết nối internet.
