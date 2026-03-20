const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    HeadingLevel,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType,
    ImageRun,
    PageBreak,
    ShadingType,
    VerticalAlign
} = require('docx');
const fs = require('fs');
const path = require('path');

// Màu sắc
const COLORS = {
    primary: '1E3A5F',      // Xanh navy đậm
    secondary: '4472C4',    // Xanh dương
    accent: '2E75B6',       // Xanh dương nhạt
    light: 'D9E2F3',        // Xanh nhạt
    text: '333333',         // Màu chữ
    white: 'FFFFFF',
    success: '70AD47',      // Xanh lá
    warning: 'FFC000',      // Vàng cam
    error: 'C00000'         // Đỏ
};

// Tạo đoạn văn bản
function createParagraph(text, options = {}) {
    return new Paragraph({
        children: [
            new TextRun({
                text: text,
                bold: options.bold || false,
                size: options.size || 24,
                color: options.color || COLORS.text,
                font: 'Times New Roman'
            })
        ],
        alignment: options.alignment || AlignmentType.LEFT,
        spacing: { after: options.afterSpacing || 200 }
    });
}

// Tạo heading
function createHeading(text, level) {
    return new Paragraph({
        text: text,
        heading: level,
        spacing: { before: 300, after: 200 }
    });
}

// Tạo bảng 2 cột (key-value)
function createTable(keyValuePairs) {
    const rows = keyValuePairs.map(([key, value]) =>
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ text: key, alignment: AlignmentType.LEFT })],
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    shading: { fill: COLORS.light, type: ShadingType.CLEAR },
                    margins: { top: 100, bottom: 100, left: 100, right: 100 }
                }),
                new TableCell({
                    children: [new Paragraph({ text: value, alignment: AlignmentType.LEFT })],
                    width: { size: 70, type: WidthType.PERCENTAGE },
                    margins: { top: 100, bottom: 100, left: 100, right: 100 }
                })
            ]
        })
    );

    return new Table({
        rows: rows,
        width: { size: 100, type: WidthType.PERCENTAGE }
    });
}

// Tạo hình ảnh placeholder
function createPlaceholderImage(widthCm, heightCm, text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: `[HÌNH ẢNH: ${text}]`,
                bold: true,
                color: COLORS.warning,
                size: 24
            })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 200 },
        border: {
            top: { style: BorderStyle.DASHED, size: 6, color: COLORS.secondary },
            bottom: { style: BorderStyle.DASHED, size: 6, color: COLORS.secondary },
            left: { style: BorderStyle.DASHED, size: 6, color: COLORS.secondary },
            right: { style: BorderStyle.DASHED, size: 6, color: COLORS.secondary }
        }
    });
}

// Tạo note box
function createNoteBox(text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: '📌 Lưu ý: ',
                bold: true,
                color: COLORS.warning
            }),
            new TextRun({
                text: text,
                color: COLORS.text
            })
        ],
        spacing: { before: 100, after: 100 },
        border: {
            left: { style: BorderStyle.SINGLE, size: 12, color: COLORS.warning }
        },
        indent: { left: 200 }
    });
}

// Tạo code block
function createCodeBlock(code) {
    return new Paragraph({
        children: [
            new TextRun({
                text: code,
                font: 'Courier New',
                size: 20,
                color: COLORS.text
            })
        ],
        spacing: { before: 100, after: 100 },
        shading: { fill: 'F5F5F5', type: ShadingType.CLEAR },
        indent: { left: 400 }
    });
}

// Tạo bullet point
function createBullet(text, level = 0) {
    return new Paragraph({
        children: [
            new TextRun({
                text: '• ' + text,
                size: 24,
                color: COLORS.text
            })
        ],
        indent: { left: 400 + (level * 200) },
        spacing: { after: 100 }
    });
}

// Tạo numbered list
function createNumberedItem(number, text) {
    return new Paragraph({
        children: [
            new TextRun({
                text: `${number}. ${text}`,
                size: 24,
                color: COLORS.text
            })
        ],
        indent: { left: 400 },
        spacing: { after: 100 }
    });
}

// Tạo step by step guide
function createStepGuide(steps) {
    return steps.map((step, index) => createNumberedItem(index + 1, step));
}

// Tạo document
async function createReport() {
    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                // ===== TRANG BÌA =====
                new Paragraph({
                    children: [
                        new TextRun({ text: '', size: 72 })
                    ],
                    spacing: { before: 2000 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'BÁO CÁO BÀI TẬP 7',
                            bold: true,
                            size: 56,
                            color: COLORS.primary
                        })
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'QUẢN LÝ INVENTORY',
                            bold: true,
                            size: 48,
                            color: COLORS.secondary
                        })
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 600 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: '─'.repeat(50), color: COLORS.light })
                    ],
                    alignment: AlignmentType.CENTER
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Môn: Nhập môn ngôn ngữ lập trình ứng dụng',
                            size: 28,
                            color: COLORS.text
                        })
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 400, after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'API RESTful với Express.js & MongoDB',
                            size: 28,
                            color: COLORS.text
                        })
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 600 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: '', size: 36 })
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `[HÌNH ẢNH TRANG BÌA - Logo/Cover Image]`,
                            bold: true,
                            color: COLORS.warning,
                            size: 24
                        })
                    ],
                    alignment: AlignmentType.CENTER
                }),

                // Page break
                new Paragraph({ children: [new PageBreak()] }),

                // ===== MỤC LỤC =====
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'MỤC LỤC',
                            bold: true,
                            size: 36,
                            color: COLORS.primary
                        })
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 400 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: '1. Giới thiệu Model Inventory', size: 24 })
                    ],
                    spacing: { after: 100 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: '2. Các chức năng API Inventory', size: 24 })
                    ],
                    spacing: { after: 100 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: '3. Hướng dẫn sử dụng Postman', size: 24 })
                    ],
                    spacing: { after: 100 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: '4. Minh họa chức năng trên Postman', size: 24 })
                    ],
                    spacing: { after: 100 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: '5. Kết luận', size: 24 })
                    ],
                    spacing: { after: 100 }
                }),

                // Page break
                new Paragraph({ children: [new PageBreak()] }),

                // ===== PHẦN 1: GIỚI THIỆU =====
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '1. GIỚI THIỆU MODEL INVENTORY',
                            bold: true,
                            size: 36,
                            color: COLORS.primary
                        })
                    ],
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 300, after: 300 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Model Inventory là phần quản lý kho hàng, theo dõi số lượng tồn kho, số lượng đặt trước (reserved) và số lượng đã bán của mỗi sản phẩm.',
                            size: 24
                        })
                    ],
                    spacing: { after: 200 }
                }),

                // Cấu trúc Model
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '1.1 Cấu trúc Model Inventory',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createTable([
                    ['Trường', 'Mô tả'],
                    ['product', 'ObjectID, ref: product, required, unique'],
                    ['stock', 'Số lượng tồn kho (min: 0)'],
                    ['reserved', 'Số lượng đặt trước (min: 0)'],
                    ['soldCount', 'Số lượng đã bán (min: 0)'],
                    ['timestamps', 'createdAt, updatedAt tự động']
                ]),
                new Paragraph({ children: [] }),

                // Mối quan hệ
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '1.2 Mối quan hệ với Product',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Mỗi khi tạo một Product mới, hệ thống sẽ tự động tạo một Inventory tương ứng với các giá trị ban đầu:',
                            size: 24
                        })
                    ],
                    spacing: { after: 100 }
                }),
                createBullet('stock = 0 (tồn kho ban đầu)'),
                createBullet('reserved = 0 (chưa có đặt trước)'),
                createBullet('soldCount = 0 (chưa có bán)'),
                new Paragraph({ children: [] }),

                // Code minh họa
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '1.3 Code tự động tạo Inventory khi tạo Product',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createCodeBlock(
`productSchema.post('save', async function(doc) {
    let inventorySchema = require('./inventory');
    let existing = await inventorySchema.findOne({ 
        product: doc._id 
    });
    if (!existing) {
        let newInventory = new inventorySchema({
            product: doc._id,
            stock: 0,
            reserved: 0,
            soldCount: 0
        });
        await newInventory.save();
    }
});`
                ),

                // Page break
                new Paragraph({ children: [new PageBreak()] }),

                // ===== PHẦN 2: CÁC CHỨC NĂNG API =====
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '2. CÁC CHỨC NĂNG API INVENTORY',
                            bold: true,
                            size: 36,
                            color: COLORS.primary
                        })
                    ],
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 300, after: 300 }
                }),

                // 2.1 Get All
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '2.1 GET ALL - Lấy tất cả Inventory',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createTable([
                    ['Phương thức', 'GET'],
                    ['URL', '/api/v1/inventory'],
                    ['Mô tả', 'Lấy danh sách tất cả inventory có join với product'],
                    ['Response', 'Array Inventory với thông tin product']
                ]),
                new Paragraph({ children: [] }),

                // 2.2 Get by ID
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '2.2 GET BY ID - Lấy Inventory theo ID',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createTable([
                    ['Phương thức', 'GET'],
                    ['URL', '/api/v1/inventory/:id'],
                    ['Mô tả', 'Lấy thông tin inventory theo ID (có join product)'],
                    ['Tham số', 'id - MongoDB ObjectID của inventory']
                ]),
                new Paragraph({ children: [] }),

                // 2.3 Add Stock
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '2.3 ADD STOCK - Thêm số lượng vào kho',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createTable([
                    ['Phương thức', 'POST'],
                    ['URL', '/api/v1/inventory/add_stock'],
                    ['Mô tả', 'Tăng stock tương ứng với quantity']
                ]),
                new Paragraph({
                    children: [new TextRun({ text: 'Body JSON:', bold: true, size: 24 })],
                    spacing: { before: 100, after: 50 }
                }),
                createCodeBlock(
`{
    "product": "<PRODUCT_ID>",
    "quantity": <NUMBER>
}`
                ),
                new Paragraph({ children: [] }),

                // 2.4 Remove Stock
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '2.4 REMOVE STOCK - Giảm số lượng khỏi kho',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createTable([
                    ['Phương thức', 'POST'],
                    ['URL', '/api/v1/inventory/remove_stock'],
                    ['Mô tả', 'Giảm stock tương ứng với quantity'],
                    ['Validation', 'Kiểm tra stock >= quantity']
                ]),
                new Paragraph({
                    children: [new TextRun({ text: 'Body JSON:', bold: true, size: 24 })],
                    spacing: { before: 100, after: 50 }
                }),
                createCodeBlock(
`{
    "product": "<PRODUCT_ID>",
    "quantity": <NUMBER>
}`
                ),
                new Paragraph({ children: [] }),

                // 2.5 Reservation
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '2.5 RESERVATION - Đặt trước sản phẩm',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createTable([
                    ['Phương thức', 'POST'],
                    ['URL', '/api/v1/inventory/reservation'],
                    ['Mô tả', 'Giảm stock, tăng reserved'],
                    ['Logic', 'stock -= quantity; reserved += quantity']
                ]),
                new Paragraph({
                    children: [new TextRun({ text: 'Body JSON:', bold: true, size: 24 })],
                    spacing: { before: 100, after: 50 }
                }),
                createCodeBlock(
`{
    "product": "<PRODUCT_ID>",
    "quantity": <NUMBER>
}`
                ),
                new Paragraph({ children: [] }),

                // 2.6 Sold
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '2.6 SOLD - Xác nhận bán sản phẩm',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createTable([
                    ['Phương thức', 'POST'],
                    ['URL', '/api/v1/inventory/sold'],
                    ['Mô tả', 'Giảm reserved, tăng soldCount'],
                    ['Logic', 'reserved -= quantity; soldCount += quantity']
                ]),
                new Paragraph({
                    children: [new TextRun({ text: 'Body JSON:', bold: true, size: 24 })],
                    spacing: { before: 100, after: 50 }
                }),
                createCodeBlock(
`{
    "product": "<PRODUCT_ID>",
    "quantity": <NUMBER>
}`
                ),

                // Page break
                new Paragraph({ children: [new PageBreak()] }),

                // ===== PHẦN 3: HƯỚNG DẪN POSTMAN =====
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '3. HƯỚNG DẪN SỬ DỤNG POSTMAN',
                            bold: true,
                            size: 36,
                            color: COLORS.primary
                        })
                    ],
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 300, after: 300 }
                }),

                // 3.1 Import Collection
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '3.1 Import Postman Collection',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                ...createStepGuide([
                    'Mở Postman',
                    'Click nút "Import" ở góc trái trên',
                    'Chọn file "NNPTUD-API.postman_collection.json"',
                    'Collection sẽ xuất hiện trong sidebar bên trái'
                ]),
                new Paragraph({ children: [] }),

                // 3.2 Tạo Category trước
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '3.2 Tạo Category (bắt buộc trước khi tạo Product)',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                ...createStepGuide([
                    'Mở folder "Categories" trong collection',
                    'Chọn request "Create Category"',
                    'Click tab "Body" và nhập JSON:',
                    'Nhấn Send và copy _id của category vừa tạo'
                ]),
                new Paragraph({
                    children: [new TextRun({ text: 'Ví dụ body:', bold: true, size: 24 })],
                    spacing: { before: 100, after: 50 }
                }),
                createCodeBlock(
`{
    "name": "Điện thoại",
    "description": "Danh mục điện thoại"
}`
                ),
                new Paragraph({ children: [] }),

                // 3.3 Tạo Product
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '3.3 Tạo Product (Inventory sẽ tự động được tạo)',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                ...createStepGuide([
                    'Mở folder "Products" trong collection',
                    'Chọn request "Create Product"',
                    'Thay <CATEGORY_ID> bằng _id đã copy ở bước trước',
                    'Nhấn Send và copy _id của product vừa tạo'
                ]),
                new Paragraph({
                    children: [new TextRun({ text: 'Ví dụ body:', bold: true, size: 24 })],
                    spacing: { before: 100, after: 50 }
                }),
                createCodeBlock(
`{
    "title": "iPhone 15 Pro Max",
    "price": 999,
    "description": "Điện thoại iPhone 15 Pro Max",
    "categoryId": "<CATEGORY_ID>",
    "images": ["https://example.com/iphone15.jpg"]
}`
                ),
                createNoteBox('Sau khi tạo Product, hệ thống sẽ tự động tạo Inventory với stock=0, reserved=0, soldCount=0'),
                new Paragraph({ children: [] }),

                // Page break
                new Paragraph({ children: [new PageBreak()] }),

                // ===== PHẦN 4: MINH HỌA CHỨC NĂNG =====
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '4. MINH HỌA CHỨC NĂNG TRÊN POSTMAN',
                            bold: true,
                            size: 36,
                            color: COLORS.primary
                        })
                    ],
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 300, after: 300 }
                }),

                // 4.1 Get All
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '4.1 Chức năng GET ALL - Lấy tất cả Inventory',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Hình ảnh minh họa:',
                            bold: true,
                            size: 24
                        })
                    ],
                    spacing: { after: 100 }
                }),
                createPlaceholderImage(15, 10, 'GET ALL - Lấy tất cả Inventory'),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '→ Sau khi tạo Product, kiểm tra Inventory đã được tạo tự động với stock=0',
                            size: 24,
                            color: COLORS.success
                        })
                    ],
                    spacing: { after: 200 }
                }),

                // 4.2 Get by ID
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '4.2 Chức năng GET BY ID - Lấy Inventory theo ID',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Hình ảnh minh họa:',
                            bold: true,
                            size: 24
                        })
                    ],
                    spacing: { after: 100 }
                }),
                createPlaceholderImage(15, 10, 'GET BY ID - Lấy Inventory theo ID'),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '→ Response trả về thông tin inventory kèm thông tin product đã join',
                            size: 24,
                            color: COLORS.success
                        })
                    ],
                    spacing: { after: 200 }
                }),

                // Page break
                new Paragraph({ children: [new PageBreak()] }),

                // 4.3 Add Stock
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '4.3 Chức năng ADD STOCK - Thêm số lượng vào kho',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Hình ảnh minh họa:',
                            bold: true,
                            size: 24
                        })
                    ],
                    spacing: { after: 100 }
                }),
                createPlaceholderImage(15, 10, 'ADD STOCK - Thêm số lượng vào kho'),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '→ Sau khi thêm, stock sẽ tăng tương ứng với quantity',
                            size: 24,
                            color: COLORS.success
                        })
                    ],
                    spacing: { after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Ví dụ: Thêm 100 sản phẩm → stock: 0 → stock: 100',
                            size: 24,
                            italics: true
                        })
                    ],
                    spacing: { after: 200 }
                }),

                // 4.4 Remove Stock
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '4.4 Chức năng REMOVE STOCK - Giảm số lượng khỏi kho',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Hình ảnh minh họa:',
                            bold: true,
                            size: 24
                        })
                    ],
                    spacing: { after: 100 }
                }),
                createPlaceholderImage(15, 10, 'REMOVE STOCK - Giảm số lượng khỏi kho'),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '→ Sau khi giảm, stock sẽ giảm tương ứng với quantity',
                            size: 24,
                            color: COLORS.success
                        })
                    ],
                    spacing: { after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Ví dụ: Giảm 10 sản phẩm → stock: 100 → stock: 90',
                            size: 24,
                            italics: true
                        })
                    ],
                    spacing: { after: 200 }
                }),
                createNoteBox('Nếu stock < quantity, hệ thống sẽ trả về lỗi "Insufficient stock"'),

                // Page break
                new Paragraph({ children: [new PageBreak()] }),

                // 4.5 Reservation
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '4.5 Chức năng RESERVATION - Đặt trước sản phẩm',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Hình ảnh minh họa:',
                            bold: true,
                            size: 24
                        })
                    ],
                    spacing: { after: 100 }
                }),
                createPlaceholderImage(15, 10, 'RESERVATION - Đặt trước sản phẩm'),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '→ Khi đặt trước: stock giảm, reserved tăng',
                            size: 24,
                            color: COLORS.success
                        })
                    ],
                    spacing: { after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Ví dụ: Đặt trước 5 sản phẩm → stock: 90 → stock: 85, reserved: 0 → reserved: 5',
                            size: 24,
                            italics: true
                        })
                    ],
                    spacing: { after: 200 }
                }),

                // 4.6 Sold
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '4.6 Chức năng SOLD - Xác nhận bán sản phẩm',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Hình ảnh minh họa:',
                            bold: true,
                            size: 24
                        })
                    ],
                    spacing: { after: 100 }
                }),
                createPlaceholderImage(15, 10, 'SOLD - Xác nhận bán sản phẩm'),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '→ Khi bán: reserved giảm, soldCount tăng',
                            size: 24,
                            color: COLORS.success
                        })
                    ],
                    spacing: { after: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Ví dụ: Bán 3 sản phẩm → reserved: 5 → reserved: 2, soldCount: 0 → soldCount: 3',
                            size: 24,
                            italics: true
                        })
                    ],
                    spacing: { after: 200 }
                }),
                createNoteBox('Nếu reserved < quantity, hệ thống sẽ trả về lỗi "Insufficient reserved quantity to sell"'),

                // Page break
                new Paragraph({ children: [new PageBreak()] }),

                // ===== PHẦN 5: TỔNG HỢP QUY TRÌNH =====
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '5. TỔNG HỢP QUY TRÌNH KHO HÀNG',
                            bold: true,
                            size: 36,
                            color: COLORS.primary
                        })
                    ],
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 300, after: 300 }
                }),

                // Sơ đồ quy trình
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '5.1 Sơ đồ luồng xử lý Inventory',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createCodeBlock(
`┌─────────────────┐
│  Tạo Product    │
│  (Auto create    │
│   Inventory=0)   │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   ADD_STOCK     │
│  stock + N      │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ REMOVE_STOCK    │
│  stock - N      │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  RESERVATION   │
│ stock-N, res+N  │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│     SOLD        │
│ res-N, sold+N   │
└─────────────────┘`
                ),
                new Paragraph({ children: [] }),

                // Bảng tổng hợp
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '5.2 Bảng tổng hợp các thay đổi',
                            bold: true,
                            size: 28,
                            color: COLORS.secondary
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                }),
                createTable([
                    ['Chức năng', 'stock', 'reserved', 'soldCount'],
                    ['ADD_STOCK', '+quantity', '0', '0'],
                    ['REMOVE_STOCK', '-quantity', '0', '0'],
                    ['RESERVATION', '-quantity', '+quantity', '0'],
                    ['SOLD', '0', '-quantity', '+quantity']
                ]),
                new Paragraph({ children: [] }),

                // Page break
                new Paragraph({ children: [new PageBreak()] }),

                // ===== PHẦN 6: KẾT LUẬN =====
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '6. KẾT LUẬN',
                            bold: true,
                            size: 36,
                            color: COLORS.primary
                        })
                    ],
                    heading: HeadingLevel.HEADING_1,
                    spacing: { before: 300, after: 300 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Bài tập 7 đã hoàn thành các yêu cầu sau:',
                            size: 24
                        })
                    ],
                    spacing: { after: 200 }
                }),
                createBullet('Model Inventory với các trường: product, stock, reserved, soldCount'),
                createBullet('Tự động tạo Inventory khi tạo Product mới (sử dụng Mongoose Middleware)'),
                createBullet('API GET ALL với join product'),
                createBullet('API GET BY ID với join product'),
                createBullet('API ADD_STOCK - tăng stock'),
                createBullet('API REMOVE_STOCK - giảm stock'),
                createBullet('API RESERVATION - giảm stock, tăng reserved'),
                createBullet('API SOLD - giảm reserved, tăng soldCount'),
                createBullet('Validation đầy đủ cho các thao tác'),
                createBullet('Postman collection đầy đủ'),
                createBullet('File Word báo cáo với hướng dẫn chi tiết'),
                new Paragraph({ children: [] }),

                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Hướng dẫn chụp ảnh màn hình Postman để điền vào báo cáo:',
                            bold: true,
                            size: 24,
                            color: COLORS.primary
                        })
                    ],
                    spacing: { before: 200, after: 100 }
                }),
                ...createStepGuide([
                    'Mở Postman và chạy từng API theo hướng dẫn ở Phần 3',
                    'Sau khi nhận được response, nhấn Cmd+Shift+4 (Mac) hoặc Win+Shift+S (Windows)',
                    'Chọn vùng cần chụp (cửa sổ Postman)',
                    'Dán vào file Word thay thế các placeholder images',
                    'Đảm bảo hình ảnh rõ ràng, có thể resize cho phù hợp'
                ]),
                new Paragraph({ children: [] }),

                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Các bước thực hiện trên Postman:',
                            bold: true,
                            size: 24,
                            color: COLORS.primary
                        })
                    ],
                    spacing: { before: 200, after: 100 }
                }),
                ...createStepGuide([
                    'Bước 1: Tạo Category → Lấy _id của category',
                    'Bước 2: Tạo Product với categoryId → Lấy _id của product (Inventory tự tạo)',
                    'Bước 3: GET ALL Inventory → Xác nhận Inventory đã tạo với stock=0',
                    'Bước 4: GET Inventory by ID → Xem chi tiết 1 inventory',
                    'Bước 5: ADD STOCK 100 → stock tăng lên 100',
                    'Bước 6: REMOVE STOCK 10 → stock giảm xuống 90',
                    'Bước 7: RESERVATION 5 → stock=85, reserved=5',
                    'Bước 8: SOLD 3 → reserved=2, soldCount=3'
                ]),
                new Paragraph({ children: [] }),

                // Footer
                new Paragraph({
                    children: [
                        new TextRun({
                            text: '─'.repeat(50),
                            color: COLORS.light
                        })
                    ],
                    alignment: AlignmentType.CENTER
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'HẾT BÁO CÁO',
                            bold: true,
                            size: 28,
                            color: COLORS.primary
                        })
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 200 }
                })
            ]
        }]
    });

    return doc;
}

// Chạy và lưu file
async function main() {
    try {
        console.log('🔄 Đang tạo báo cáo Word...');
        const doc = await createReport();
        const buffer = await Packer.toBuffer(doc);
        
        const outputPath = path.join(__dirname, 'BaoCao_BT7_Inventory.docx');
        fs.writeFileSync(outputPath, buffer);
        
        console.log(`✅ Đã tạo file: ${outputPath}`);
        console.log('📝 Vui lòng mở file Word và thay thế các placeholder images bằng ảnh chụp màn hình từ Postman');
    } catch (error) {
        console.error('❌ Lỗi khi tạo file:', error);
    }
}

main();
