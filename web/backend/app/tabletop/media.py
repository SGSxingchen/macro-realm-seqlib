"""Accept raster images only. Decode/re-encode so metadata and active content are stripped."""
from io import BytesIO
import warnings
from PIL import Image, UnidentifiedImageError
from fastapi import HTTPException

def sanitize(data: bytes):
    if len(data) > 5 * 1024 * 1024:
        raise HTTPException(413, '图片不得超过 5MB')
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('error', Image.DecompressionBombWarning)
            with Image.open(BytesIO(data)) as image:
                if image.format not in ('PNG','JPEG','WEBP') or image.width*image.height > 16_000_000:
                    raise ValueError('不支持的图片或尺寸过大')
                if image.width > 8192 or image.height > 8192:
                    raise ValueError('图片边长不得超过8192像素')
                image.load()
                image = image.convert('RGBA')
                result = BytesIO()
                image.save(result, format='PNG')
                raw=result.getvalue()
                if len(raw)>5*1024*1024: raise ValueError('解码后的图片超过5MB，请降低分辨率')
                return raw, image.width, image.height
    except (UnidentifiedImageError, OSError, ValueError, Image.DecompressionBombError, Image.DecompressionBombWarning) as exc:
        raise HTTPException(400,'无法使用此图片：仅支持安全的 PNG/JPEG/WebP（最大1600万像素）') from exc
