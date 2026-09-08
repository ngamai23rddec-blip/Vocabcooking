# VocaBloom

Web học từ vựng tiếng Anh — React + Vite. Có thể chạy **có đăng nhập** (đồng bộ dữ liệu theo tài khoản, qua Supabase) hoặc **không đăng nhập** (dữ liệu lưu ngay trên trình duyệt) — mặc định là không đăng nhập, bạn bật tính năng đăng nhập bất cứ lúc nào theo hướng dẫn ở Phần 3.

## 1. Chạy thử trên máy (local)

Cần cài **Node.js** (bản 18 trở lên) trước: https://nodejs.org

```bash
cd vocabloom-site
npm install
npm run dev
```

Mở trình duyệt vào địa chỉ hiện lên trong terminal (thường là `http://localhost:5173`).

## 2. Đưa lên internet — cách đáng tin cậy nhất: Vercel CLI

Việc kéo-thả file lên GitHub rất dễ bị lỗi (thiếu file, sai cấu trúc thư mục, lỡ tay tải nguyên file .zip lên...). Cách chắc chắn nhất, ít bước nhất, và đúng 100% cấu trúc là deploy thẳng từ máy bạn lên Vercel bằng dòng lệnh — không cần qua GitHub.

Trong terminal, tại đúng thư mục `vocabloom-site` (đã chạy `npm install` ở bước 1):

```bash
npx vercel login
```
→ Làm theo hướng dẫn (nhập email, Vercel gửi link xác nhận, bấm vào link đó).

```bash
npx vercel --prod
```
→ Trả lời các câu hỏi:
- "Set up and deploy?" → gõ `y`
- "Which scope?" → chọn tài khoản của bạn (Enter)
- "Link to existing project?" → gõ `N`
- "What's your project's name?" → Enter để giữ mặc định (`vocabloom-site`)
- "In which directory is your code located?" → Enter để giữ mặc định (`./`)
- Các câu còn lại → cứ Enter để giữ mặc định

Đợi khoảng 30–60 giây, Vercel sẽ in ra một đường link dạng `https://vocabloom-site-xxxx.vercel.app` — đó là website thật của bạn.

**Muốn cập nhật sau này:** mỗi khi sửa code xong, chạy lại đúng lệnh `npx vercel --prod` trong thư mục đó là web tự cập nhật.

### Nếu vẫn muốn dùng GitHub + Vercel (tự động deploy mỗi khi sửa code)
Vào repo GitHub, kiểm tra chắc chắn các file `package.json`, `index.html`, `vite.config.js`, `vercel.json` và thư mục `src/` (chứa `App.jsx`, `main.jsx`) nằm **ngay tại gốc repo** — không bị lồng thêm một cấp thư mục `vocabloom-site/` nào nữa, và không có file `.zip` nào còn sót lại trong repo. Nếu cấu trúc đúng, import vào Vercel như bình thường (Add New → Project → chọn repo → Deploy).

## 3. Bật tính năng đăng nhập tài khoản (tuỳ chọn)

Mặc định web **không yêu cầu đăng nhập** — ai vào cũng dùng được ngay, dữ liệu lưu trên trình duyệt của họ, và Vercel build sẽ không bao giờ lỗi vì thiếu cấu hình. Nếu bạn muốn có tài khoản thật (đăng ký/đăng nhập bằng email, dữ liệu đồng bộ theo tài khoản trên mọi thiết bị), làm theo các bước sau — làm lúc nào cũng được, không bắt buộc phải làm ngay.

### Bước 1 — Tạo project Supabase (miễn phí)
1. Vào https://supabase.com → **Start your project** → đăng nhập bằng GitHub.
2. Bấm **New project**, đặt tên (ví dụ `vocabloom`), đặt một mật khẩu database bất kỳ (nhớ lưu lại phòng khi cần), chọn Region gần bạn nhất (ví dụ Singapore), bấm **Create new project**. Đợi khoảng 1–2 phút để Supabase khởi tạo.

### Bước 2 — Tạo bảng lưu dữ liệu người dùng
1. Ở menu bên trái, bấm biểu tượng **SQL Editor** (`>_`).
2. Bấm **New query**, dán đoạn SQL dưới đây vào, rồi bấm **Run**:

```sql
create table if not exists vocabloom_data (
  user_id uuid primary key references auth.users(id) on delete cascade,
  payload jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

alter table vocabloom_data enable row level security;

create policy "Users can view their own data"
  on vocabloom_data for select
  using (auth.uid() = user_id);

create policy "Users can insert their own data"
  on vocabloom_data for insert
  with check (auth.uid() = user_id);

create policy "Users can update their own data"
  on vocabloom_data for update
  using (auth.uid() = user_id);
```

Đoạn này tạo một bảng lưu dữ liệu (từ vựng, tiến độ...) và bật bảo mật để mỗi người chỉ đọc/ghi được đúng dữ liệu của chính họ.

### Bước 3 — Lấy khoá kết nối
1. Vào **Project Settings** (biểu tượng bánh răng) → **API**.
2. Copy 2 giá trị: **Project URL** và khoá **anon public**.

### Bước 4 — (Khuyên dùng để test nhanh) Tắt yêu cầu xác nhận email
Mặc định Supabase bắt người dùng bấm vào link trong email mới đăng nhập được lần đầu. Muốn đăng ký xong dùng luôn cho nhanh:
1. Vào **Authentication** → **Providers** → **Email**.
2. Tắt **Confirm email**, bấm **Save**.
(Có thể bật lại sau này khi web đã có người dùng thật, để an toàn hơn.)

### Bước 5 — Thêm khoá vào Vercel
1. Vào project của bạn trên vercel.com → **Settings** → **Environment Variables**.
2. Thêm 2 biến:
   - `VITE_SUPABASE_URL` = Project URL vừa copy
   - `VITE_SUPABASE_ANON_KEY` = khoá anon public vừa copy
3. Bấm **Save**.

### Bước 6 — Deploy lại
- Nếu dùng Vercel CLI: chạy lại `npx vercel --prod` trong thư mục `vocabloom-site`.
- Nếu dùng GitHub: vào tab **Deployments** trên Vercel → bấm **Redeploy** ở bản mới nhất.

Sau bước này, mở lại link web sẽ thấy màn hình đăng nhập/đăng ký xuất hiện tự động.

## 4. Lưu ý quan trọng

- **Không đăng nhập:** dữ liệu lưu trong `localStorage` của trình duyệt — lưu lại khi refresh, nhưng không đồng bộ giữa các thiết bị, và mất nếu người dùng xoá dữ liệu duyệt web.
- **Có đăng nhập (Supabase):** dữ liệu lưu theo tài khoản trên server của Supabase — mở từ điện thoại, máy tính, trình duyệt nào cũng thấy đúng dữ liệu đó, miễn là đăng nhập cùng một tài khoản.
- **Phát âm** dùng công nghệ có sẵn của trình duyệt (Web Speech API), không cần backend.
- **Import Excel** chạy hoàn toàn phía trình duyệt, không upload file lên server nào.
- File `vercel.json` đã cấu hình sẵn build command/output directory rõ ràng, tránh việc Vercel đoán sai cấu hình.
