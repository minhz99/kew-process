KEW6315_REF_WIDTH = 240
KEW6315_REF_HEIGHT = 256


def make_grid(ids, x_rights, y_bot, bg):
    return [{"id": id_, "x": x, "y": y_bot, "bg": bg} for id_, x in zip(ids, x_rights)]


def _map_sd140(overlay):
    if overlay["id"] in ["PF1", "PF2", "PF3"]:
        overlay["w_clear"] = 55
    return overlay


SCREENS = [
    {
        "id": "SD140",
        "overlays": list(map(_map_sd140,
            make_grid(["V1", "V2", "V3"], [94, 158, 222], 54, "w") +
            make_grid(["A1", "A2", "A3"], [94, 158, 222], 70, "g") +
            make_grid(["P1", "P2", "P3"], [94, 158, 222], 86, "w") +
            make_grid(["Q1", "Q2", "Q3"], [94, 158, 222], 102, "g") +
            make_grid(["S1", "S2", "S3"], [94, 158, 222], 118, "w") +
            make_grid(["PF1", "PF2", "PF3"], [94, 158, 222], 134, "g") +
            [
                {"id": "P", "x": 94, "y": 153, "bg": "w"},
                {"id": "freq", "alias": "f", "x": 222, "y": 153, "bg": "w"},
                {"id": "Q", "x": 94, "y": 169, "bg": "g"},
                {"id": "S", "x": 94, "y": 185, "bg": "w"},
                {"id": "PF", "x": 94, "y": 201, "bg": "g", "w_clear": 55},
                {"id": "An", "x": 222, "y": 201, "bg": "g"}
            ]
        ))
    },
    {
        "id": "SD141",
        "overlays": [
            {"id": "V1", "x": 63, "y": 36, "bg": "w", "w_clear": 45},
            {"id": "Vdeg1", "x": 121, "y": 36, "bg": "w", "w_clear": 45},
            {"id": "V2", "x": 63, "y": 52, "bg": "g", "w_clear": 45},
            {"id": "Vdeg2", "x": 121, "y": 52, "bg": "g", "w_clear": 45},
            {"id": "V3", "x": 63, "y": 68, "bg": "w", "w_clear": 45},
            {"id": "Vdeg3", "x": 121, "y": 68, "bg": "w", "w_clear": 45},
            {"id": "A1", "x": 63, "y": 87, "bg": "w", "w_clear": 45},
            {"id": "Adeg1", "x": 121, "y": 87, "bg": "w", "w_clear": 45},
            {"id": "A2", "x": 63, "y": 103, "bg": "g", "w_clear": 45},
            {"id": "Adeg2", "x": 121, "y": 103, "bg": "g", "w_clear": 45},
            {"id": "A3", "x": 63, "y": 119, "bg": "w", "w_clear": 45},
            {"id": "Adeg3", "x": 121, "y": 119, "bg": "w", "w_clear": 45},
            {"id": "freq", "alias": "f", "x": 83, "y": 154, "bg": "w", "w_clear": 45},
            {"id": "V_unb", "alias": "V%", "x": 83, "y": 189, "bg": "g", "w_clear": 45},
            {"id": "A_unb", "alias": "A%", "x": 83, "y": 205, "bg": "w", "w_clear": 45}
        ]
    },
    {
        "id": "SD142",
        "overlays": make_grid(["V1", "V2", "V3"], [76, 136, 196], 47, "w") +
        make_grid(["A1", "A2", "A3"], [76, 136, 196], 63, "g")
    },
    {
        "id": "SD143",
        "overlays": make_grid(["V1", "V2", "V3"], [76, 136, 196], 47, "w") +
        make_grid(["A1", "A2", "A3"], [76, 136, 196], 63, "g")
    },
    {
        "id": "SD144",
        "overlays": make_grid(["V1", "V2", "V3"], [76, 136, 196], 47, "w") +
        make_grid(["THDV1", "THDV2", "THDV3"], [76, 136, 196], 63, "g")
    },
    {
        "id": "SD145",
        "overlays": make_grid(["A1", "A2", "A3"], [76, 136, 196], 47, "w") +
        make_grid(["THDA1", "THDA2", "THDA3"], [76, 136, 196], 63, "g")
    }
]


SCREEN_BY_ID = {screen["id"]: screen for screen in SCREENS}
SCREEN_BY_INDEX = {index: screen for index, screen in enumerate(SCREENS)}
