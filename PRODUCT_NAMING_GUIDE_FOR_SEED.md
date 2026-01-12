# 📝 Hướng Dẫn Đặt Tên Sản Phẩm Cho Seed Data

## 🎯 Nguyên Tắc Cơ Bản

### 1. Format Tên Sản Phẩm (name)

```
[Loại sản phẩm] [Fit/Style] - [Tên Design]
```

**Ví dụ:**
```
✅ "Áo Thun Relaxed Fit - Sushi Meow"
✅ "Áo Khoác Bomber - Street Style"
✅ "Quần Jean Slim Fit - Dark Blue"
✅ "Áo Sơ Mi Dài Tay - Classic White"
✅ "Áo Hoodie Oversized - Minimalist Black"
```

**Loại sản phẩm phổ biến:**
- Áo Thun (T-Shirt)
- Áo Sơ Mi (Shirt)
- Áo Khoác (Jacket)
- Áo Hoodie (Hoodie)
- Áo Polo (Polo)
- Áo Len (Sweater)
- Quần Jean (Jeans)
- Quần Kaki (Khakis)
- Quần Short (Shorts)

**Fit/Style:**
- Slim Fit, Regular Fit, Relaxed Fit, Oversized
- Dài Tay, Ngắn Tay, Ba Lỗ
- Bomber, Denim, Cardigan, Zip-up

---

### 2. Mô Tả Sản Phẩm (description)

**Cấu trúc:**
```
[Giới thiệu ngắn] + [Đặc điểm nổi bật] + [Phù hợp với] + [Keywords]
```

**Template:**
```
[Tên sản phẩm] với [đặc điểm thiết kế]. 
Chất liệu [material] [đặc tính: mềm mại, thoáng mát...]. 
Phù hợp cho [phong cách/dịp]. 
[Thêm keywords tìm kiếm]
```

**Ví dụ:**
```
Áo thun Relaxed Fit với họa tiết Sushi Meow đáng yêu, kết hợp giữa văn hóa ẩm thực Nhật Bản và hình tượng mèo kawaii. 
Chất liệu cotton 100% mềm mại, thoáng mát, thấm hút mồ hôi tốt. 
Phù hợp cho phong cách casual hàng ngày, đi chơi cuối tuần, đi cafe với bạn bè.
Keywords: áo thun, áo mèo, áo đồ ăn, áo nhật bản, áo cute, áo kawaii, áo họa tiết.
```

---

### 3. Keywords Quan Trọng Cần Có

#### A. Loại Sản Phẩm (BẮT BUỘC)
```
Tiếng Việt + Tiếng Anh + Từ đồng nghĩa
```
- Áo thun, t-shirt, áo phông
- Áo khoác, jacket, khoác
- Quần jean, jeans, jean
- Áo sơ mi, shirt, sơ mi

#### B. Họa Tiết / Theme
```
Nếu có họa tiết đặc biệt → nhắc trong description
```
- **Động vật:** mèo, chó, gấu, hổ, rồng (cat, dog, bear, tiger, dragon)
- **Thực vật:** hoa, lá, cây (flower, floral, leaf, botanical)
- **Đồ vật:** xe, máy bay, nhạc cụ (car, plane, music)
- **Trừu tượng:** kẻ sọc, chấm bi, trơn (stripe, polka dot, solid)
- **Văn hóa:** nhật bản, hàn quốc, vintage (japanese, korean, vintage)
- **Phong cách:** minimalist, streetwear, retro

#### C. Màu Sắc Chính
```
Nhắc màu chính trong tên hoặc description
```
- Đen (black), Trắng (white), Xanh (blue), Đỏ (red)
- Nâu (brown), Xám (gray), Be (beige), Hồng (pink)

#### D. Phong Cách / Style
```
Giúp user tìm theo phong cách ăn mặc
```
- **Casual:** thường ngày, thoải mái, đi chơi
- **Formal:** công sở, lịch sự, sang trọng
- **Sporty:** thể thao, năng động, gym
- **Streetwear:** đường phố, hip hop, urban
- **Vintage:** retro, cổ điển, hoài cổ
- **Minimalist:** tối giản, đơn giản, tinh tế

#### E. Dịp Sử Dụng
```
Giúp chatbot recommend theo context
```
- Hàng ngày, đi làm, đi học
- Đi chơi, đi cafe, hẹn hò
- Dự tiệc, sự kiện, đám cưới
- Đi biển, du lịch, picnic
- Thể thao, tập gym, chạy bộ

---

## 📋 Template Cho Từng Loại Sản Phẩm

### Áo Thun (T-Shirt)

**Name:**
```
Áo Thun [Fit] - [Design Name]
```

**Description Template:**
```
Áo thun [fit] với [đặc điểm design: họa tiết X, in chữ Y, màu Z...]. 
Chất liệu cotton [đặc tính]. 
Phù hợp [phong cách + dịp].
Keywords: áo thun, t-shirt, áo phông, [theme], [màu sắc], [style].
```

**Ví dụ:**
```json
{
  "name": "Áo Thun Oversized - Meow Gang",
  "description": "Áo thun oversized với họa tiết Meow Gang gồm 4 chú mèo đáng yêu. Chất liệu cotton 100% mềm mại, form rộng thoải mái. Phù hợp phong cách streetwear, đi chơi cuối tuần. Keywords: áo thun, áo mèo, áo oversized, áo cute, áo đen."
}
```

---

### Áo Sơ Mi (Shirt)

**Name:**
```
Áo Sơ Mi [Dài/Ngắn Tay] - [Design/Color]
```

**Description Template:**
```
Áo sơ mi [tay dài/ngắn] [đặc điểm: kẻ sọc, trơn, họa tiết...]. 
Chất liệu [cotton/linen/blend]. 
Phù hợp [công sở, dự tiệc, casual...].
Keywords: áo sơ mi, shirt, [màu], [pattern].
```

**Ví dụ:**
```json
{
  "name": "Áo Sơ Mi Dài Tay - Classic White",
  "description": "Áo sơ mi dài tay trắng basic, thiết kế tối giản sang trọng. Chất liệu cotton pha linen thoáng mát. Phù hợp đi làm, dự sự kiện, mặc đi học. Keywords: áo sơ mi, shirt, áo trắng, áo công sở, áo formal."
}
```

---

### Áo Khoác (Jacket)

**Name:**
```
Áo Khoác [Type] - [Design/Color]
```

**Description Template:**
```
Áo khoác [bomber/denim/windbreaker...] [đặc điểm thiết kế]. 
Chất liệu [material] [ưu điểm: chống gió, giữ ấm...]. 
Phù hợp [mùa, dịp].
Keywords: áo khoác, jacket, [type], [style].
```

**Ví dụ:**
```json
{
  "name": "Áo Khoác Bomber - Pilot Black",
  "description": "Áo khoác bomber style phi công, màu đen cá tính. Chất liệu polyester chống gió nhẹ, có lớp lót giữ ấm. Phù hợp mùa thu đông, phong cách streetwear. Keywords: áo khoác, bomber jacket, áo đen, áo streetwear, áo pilot."
}
```

---

### Quần (Pants/Jeans)

**Name:**
```
Quần [Type] [Fit] - [Design/Color]
```

**Description Template:**
```
Quần [jean/kaki/jogger...] [slim/regular/straight...] [đặc điểm]. 
Chất liệu [denim/cotton...]. 
Phù hợp [style, dịp].
Keywords: quần [type], [fit], [màu].
```

**Ví dụ:**
```json
{
  "name": "Quần Jean Slim Fit - Dark Blue",
  "description": "Quần jean slim fit màu xanh đậm, co giãn nhẹ thoải mái. Chất liệu denim cao cấp, giữ form tốt. Phù hợp đi làm, đi chơi, phong cách smart casual. Keywords: quần jean, jeans, slim fit, quần xanh, quần đi làm."
}
```

---

## 🎨 Ví Dụ Theo Theme

### Theme: Động Vật (Animals)

```
"Áo Thun Relaxed Fit - Cat Cafe"
→ Keywords: áo thun, áo mèo, áo cafe, áo cute, áo kawaii

"Áo Khoác Hoodie - Bear Hug"
→ Keywords: áo hoodie, áo gấu, áo cute, áo ấm

"Áo Thun Oversized - Tiger Street"
→ Keywords: áo thun, áo hổ, áo streetwear, áo oversized
```

### Theme: Đồ Ăn (Food)

```
"Áo Thun Regular Fit - Sushi Love"
→ Keywords: áo thun, áo sushi, áo đồ ăn, áo nhật bản

"Áo Thun Crop Top - Pizza Party"
→ Keywords: áo thun, áo pizza, áo đồ ăn, áo crop top, áo đi chơi
```

### Theme: Thiên Nhiên (Nature)

```
"Áo Sơ Mi Ngắn Tay - Tropical Vibes"
→ Keywords: áo sơ mi, áo hoa, áo nhiệt đới, áo đi biển

"Áo Thun Regular Fit - Mountain Sunset"
→ Keywords: áo thun, áo núi, áo thiên nhiên, áo phong cảnh
```

### Theme: Trừu Tượng (Abstract/Pattern)

```
"Áo Thun Oversized - Minimalist Black"
→ Keywords: áo thun, áo đen, áo minimalist, áo đơn giản, áo basic

"Áo Sơ Mi Dài Tay - Blue Stripes"
→ Keywords: áo sơ mi, áo kẻ sọc, áo xanh, áo công sở
```

---

## ✅ Checklist Khi Seed Data

- [ ] **Name:** Có loại sản phẩm tiếng Việt (Áo Thun, Áo Khoác...)
- [ ] **Name:** Có fit/style nếu phù hợp (Slim Fit, Oversized...)
- [ ] **Name:** Có tên design độc đáo
- [ ] **Description:** Mô tả đầy đủ đặc điểm sản phẩm
- [ ] **Description:** Có chất liệu và ưu điểm
- [ ] **Description:** Có phong cách/dịp sử dụng
- [ ] **Description:** Có keywords ở cuối (cả tiếng Việt + Anh)
- [ ] **Keywords:** Bao gồm loại sản phẩm, theme, màu sắc, style

---

## 💡 Tips

1. **Luôn dùng tiếng Việt cho loại sản phẩm** trong name → User Việt hay tìm "áo thun" hơn "t-shirt"

2. **Keywords phải bao gồm cả từ đồng nghĩa:**
   - Áo thun = t-shirt = áo phông
   - Áo khoác = jacket = khoác

3. **Theme/họa tiết quan trọng cho search:**
   - User tìm "áo mèo" → phải match được
   - User tìm "áo hoa" → phải match được

4. **Màu sắc nên xuất hiện trong name hoặc description:**
   - "- Dark Blue" trong name
   - "màu đen" trong description

5. **Phong cách giúp recommend theo context:**
   - "casual" → đi chơi hàng ngày
   - "formal" → đi làm, dự tiệc
   - "streetwear" → phong cách đường phố

6. **Description càng chi tiết → search càng tốt**
   - Nhưng giữ cho tự nhiên, không spam keywords
