"""
kew6315_layout.py
=================
Layout definitions cho ảnh KEW6315 thực tế từ máy đo.
Kích thước ảnh thực tế: 320x240 (width x height)
"""

KEW6315_REF_WIDTH = 320
KEW6315_REF_HEIGHT = 240


def make_grid(ids, x_rights, y_bot, bg):
    """Tạo overlay grid cho 3 cột."""
    return [{"id": id_, "x": x, "y": y_bot, "bg": bg} for id_, x in zip(ids, x_rights)]


def _map_sd140(overlay):
    """SD140: PF fields cần width lớn hơn."""
    if overlay["id"] in ["PF1", "PF2", "PF3", "PF"]:
        overlay["w_clear"] = 70
    return overlay


SCREENS = [
    {
        # Màn hình CHÍNH (SD140): V, A, P, Q, S, PF
        # Phân tích từ PS-SD785.BMP
        "id": "SD140",
        "overlays": list(map(_map_sd140,
            make_grid(["V1", "V2", "V3"], [93, 157, 221], 50, "w") +
            make_grid(["A1", "A2", "A3"], [93, 157, 221], 66, "g") +
            make_grid(["P1", "P2", "P3"], [93, 157, 221], 82, "w") +
            make_grid(["Q1", "Q2", "Q3"], [93, 157, 221], 98, "g") +
            make_grid(["S1", "S2", "S3"], [93, 157, 221], 114, "w") +
            make_grid(["PF1", "PF2", "PF3"], [93, 157, 221], 130, "g") +
            [
                {"id": "P", "x": 93, "y": 146, "bg": "w"},
                {"id": "freq", "alias": "f", "x": 221, "y": 146, "bg": "w"},
                {"id": "Q", "x": 93, "y": 156, "bg": "g"},
                {"id": "S", "x": 93, "y": 166, "bg": "w"},
                {"id": "PF", "x": 93, "y": 176, "bg": "g", "w_clear": 70},
                {"id": "An", "x": 221, "y": 176, "bg": "g"}
            ]
        ))
    },
    {
        # Màn hình GÓC PHA (SD141): V/A phase angles + V_unb/A_unb
        "id": "SD141",
        "overlays": [
            {"id": "V1", "x": 83, "y": 48, "bg": "w", "w_clear": 60},
            {"id": "Vdeg1", "x": 163, "y": 48, "bg": "w", "w_clear": 60},
            {"id": "V2", "x": 83, "y": 64, "bg": "g", "w_clear": 60},
            {"id": "Vdeg2", "x": 163, "y": 64, "bg": "g", "w_clear": 60},
            {"id": "V3", "x": 83, "y": 80, "bg": "w", "w_clear": 60},
            {"id": "Vdeg3", "x": 163, "y": 80, "bg": "w", "w_clear": 60},
            {"id": "A1", "x": 83, "y": 99, "bg": "w", "w_clear": 60},
            {"id": "Adeg1", "x": 163, "y": 99, "bg": "w", "w_clear": 60},
            {"id": "A2", "x": 83, "y": 115, "bg": "g", "w_clear": 60},
            {"id": "Adeg2", "x": 163, "y": 115, "bg": "g", "w_clear": 60},
            {"id": "A3", "x": 83, "y": 131, "bg": "w", "w_clear": 60},
            {"id": "Adeg3", "x": 163, "y": 131, "bg": "w", "w_clear": 60},
            {"id": "freq", "alias": "f", "x": 110, "y": 165, "bg": "w", "w_clear": 60},
            {"id": "V_unb", "alias": "V%", "x": 110, "y": 195, "bg": "g", "w_clear": 60},
            {"id": "A_unb", "alias": "A%", "x": 110, "y": 210, "bg": "w", "w_clear": 60}
        ]
    },
    {
        # Màn hình VECTOR 1 (SD142): V và A
        "id": "SD142",
        "overlays": make_grid(["V1", "V2", "V3"], [100, 180, 260], 62, "w") +
        make_grid(["A1", "A2", "A3"], [100, 180, 260], 84, "g")
    },
    {
        # Màn hình VECTOR 2 (SD143): V và A (tương tự SD142)
        "id": "SD143",
        "overlays": make_grid(["V1", "V2", "V3"], [100, 180, 260], 62, "w") +
        make_grid(["A1", "A2", "A3"], [100, 180, 260], 84, "g")
    },
    {
        # Màn hình THD ĐIỆN ÁP (SD144): V + THDV
        "id": "SD144",
        "overlays": make_grid(["V1", "V2", "V3"], [100, 180, 260], 62, "w") +
        make_grid(["THDV1", "THDV2", "THDV3"], [100, 180, 260], 84, "g")
    },
    {
        # Màn hình THD DÒNG ĐIỆN (SD145): A + THDA
        "id": "SD145",
        "overlays": make_grid(["A1", "A2", "A3"], [100, 180, 260], 62, "w") +
        make_grid(["THDA1", "THDA2", "THDA3"], [100, 180, 260], 84, "g")
    }
]


SCREEN_BY_INDEX = {index: screen for index, screen in enumerate(SCREENS)}
