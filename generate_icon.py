"""초록 나무 아이콘 생성 스크립트 (한 번만 실행)"""
from PIL import Image, ImageDraw


def draw_tree(img: Image.Image, size: int) -> None:
    """이미지 위에 초록 나무를 그린다."""
    d = ImageDraw.Draw(img)
    w = h = size

    # 배경: 부드러운 크림색 원 (아이콘 배경)
    margin = max(1, size // 32)
    d.ellipse(
        [margin, margin, w - margin, h - margin],
        fill=(250, 248, 220, 255),  # 크림색
        outline=(80, 60, 20, 255),
        width=max(1, size // 48),
    )

    # 줄기 (갈색 사각형)
    trunk_w = max(2, w // 7)
    trunk_h = max(3, h // 5)
    trunk_x = (w - trunk_w) // 2
    trunk_y = int(h * 0.68)
    d.rectangle(
        [trunk_x, trunk_y, trunk_x + trunk_w, trunk_y + trunk_h],
        fill=(101, 67, 33, 255),  # 진한 갈색
    )

    # 잎사귀: 3단 삼각형 (진한 녹색 → 밝은 녹색 그라데이션 효과)
    cx = w // 2

    # 맨 아래 (가장 큰 잎, 어두운 녹색)
    y_bot = int(h * 0.72)
    y_top1 = int(h * 0.45)
    tri1 = [(cx, y_top1), (int(w * 0.12), y_bot), (int(w * 0.88), y_bot)]
    d.polygon(tri1, fill=(22, 101, 52, 255))  # 진한 숲 녹색

    # 가운데 잎 (중간 녹색)
    y_bot2 = int(h * 0.56)
    y_top2 = int(h * 0.26)
    tri2 = [(cx, y_top2), (int(w * 0.18), y_bot2), (int(w * 0.82), y_bot2)]
    d.polygon(tri2, fill=(46, 139, 87, 255))  # 바다 녹색

    # 맨 위 잎 (밝은 녹색)
    y_bot3 = int(h * 0.38)
    y_top3 = int(h * 0.08)
    tri3 = [(cx, y_top3), (int(w * 0.25), y_bot3), (int(w * 0.75), y_bot3)]
    d.polygon(tri3, fill=(76, 175, 80, 255))  # 밝은 초록


def create_icon(output_path: str) -> None:
    # 256x256에서 그린 뒤 PIL이 자동으로 모든 크기로 리사이즈하도록 한다
    base_size = 256
    img = Image.new("RGBA", (base_size, base_size), (0, 0, 0, 0))
    draw_tree(img, base_size)

    sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    img.save(output_path, format="ICO", sizes=sizes)
    print(f"아이콘 생성 완료: {output_path}")


if __name__ == "__main__":
    import os
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "icon.ico")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    create_icon(out)
