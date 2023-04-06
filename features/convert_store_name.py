if 1 == 1:
    import sys
    import warnings
    import os

    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
    warnings.simplefilter("ignore", UserWarning)
    sys.coinit_flags = 2


class StoreNameConverter:
    def __init__(self):
        print(f"상점 이름 변환")

    # 엑셀에 적혀있는 상점의 이름을 이지어드민에 있는 이름으로 변환해야 합니다.
    def convert_store_name(self, store_name: str):
        print(f"{store_name}")

        if store_name == "11번가":
            store_name = "11번가"
        elif store_name == "지마켓":
            store_name = "G마켓"
        # elif store_name == "":
        #     store_name = "R2O"
        # elif store_name == "":
        #     store_name = "개별주문"
        # elif store_name == "":
        #     store_name = "그립"
        elif store_name == "브랜디":
            store_name = "브랜디"
        elif store_name == "브리치":
            store_name = "브리치"
        # elif store_name == "":
        #     store_name = "쇼핑엔티"
        # elif store_name == "":
        #     store_name = "스토어팜"
        # elif store_name == "":
        #     store_name = "에이블리"
        elif store_name == "위메프":
            store_name = "위메이크프라이스 2.0"
        # elif store_name == "":
        #     store_name = "직진배송"
        elif store_name == "카카오":
            store_name = "카카오톡스토어"
        elif store_name == "자사몰(까페24기준)" or store_name == "카페24" or store_name == "까페24":
            store_name = "카페24"
        elif store_name == "쿠팡":
            store_name = "쿠팡(자동)"
        # elif store_name == "":
        #     store_name = "쿠팡로켓배송(송장)"
        elif store_name == "티몬":
            store_name = "티몬(자동)"
        # elif store_name == "":
        #     store_name = "티켓몬스터(자동) #4"

        return store_name
