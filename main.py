import os
import zipfile
import io
from PIL import Image

def compress_xlsx(input_xlsx, resize_ratio=0.5, quality=75):
    """
    压缩 Excel 文件中的图片并重新打包
    :param input_xlsx: 原始 xlsx 文件路径
    :param resize_ratio: 图片缩放比例 (0.5 = 宽高各减半)
    :param quality: 图片保存质量 (1-95)
    """
    if not os.path.isfile(input_xlsx):
        print(f"文件不存在: {input_xlsx}")
        return

    if not input_xlsx.lower().endswith('.xlsx'):
        print("仅支持 .xlsx 文件")
        return

    output_xlsx = input_xlsx[:-5] + "_compressed.xlsx"
    temp_images = 0
    saved_size = 0

    with zipfile.ZipFile(input_xlsx, 'r') as zin:
        with zipfile.ZipFile(output_xlsx, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                # 只处理图片
                if item.filename.startswith('xl/media/'):
                    try:
                        img = Image.open(io.BytesIO(data))
                        w, h = img.size
                        img = img.resize((int(w * resize_ratio), int(h * resize_ratio)))
                        buf = io.BytesIO()
                        img_format = img.format if img.format else "JPEG"
                        img.save(buf, format=img_format, optimize=True, quality=quality)
                        new_data = buf.getvalue()
                        zout.writestr(item.filename, new_data)
                        temp_images += 1
                        saved_size += len(data) - len(new_data)
                        continue
                    except Exception as e:
                        print(f"图片处理失败 {item.filename}: {e}")
                # 非图片文件直接写入
                zout.writestr(item, data)

    print(f"已生成压缩文件: {output_xlsx}")
    print(f"压缩图片数量: {temp_images}")
    print(f"节省空间约: {saved_size/1024:.2f} KB")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="压缩 xlsx 文件中的图片")
    parser.add_argument("input", help="输入的 xlsx 文件路径")
    parser.add_argument("--resize", type=float, default=0.8, help="缩放比例 (默认 0.8)")
    parser.add_argument("--quality", type=int, default=85, help="图片质量 (默认 85)")
    args = parser.parse_args()

    compress_xlsx(args.input, resize_ratio=args.resize, quality=args.quality)

