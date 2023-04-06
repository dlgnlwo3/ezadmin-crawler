class StoreDetailDto:
    def __init__(self):
        self.__store_name = ""

        self.__tot_products = ""  # 주문수량

        self.__tot_amount = ""  # 주문금액

        self.__org_price = ""  # 상품원가

        self.__cancel_total_data_product_sum = ""  # 취소수량

        self.__cancel_total_data_order_sum_amount = ""  # 취소금액

        self.__refund_total_data_product_sum = ""  # 반품수량

        self.__refund_total_data_order_sum_amount = ""  # 반품금액

    @property
    def store_name(self):  # getter
        return self.__store_name

    @store_name.setter
    def store_name(self, value: str):  # setter
        self.__store_name = value

    @property
    def tot_products(self):  # getter
        return self.__tot_products

    @tot_products.setter
    def tot_products(self, value: str):  # setter
        self.__tot_products = value

    @property
    def tot_amount(self):  # getter
        return self.__tot_amount

    @tot_amount.setter
    def tot_amount(self, value: str):  # setter
        self.__tot_amount = value

    @property
    def org_price(self):  # getter
        return self.__org_price

    @org_price.setter
    def org_price(self, value: str):  # setter
        self.__org_price = value

    @property
    def cancel_total_data_product_sum(self):  # getter
        return self.__cancel_total_data_product_sum

    @cancel_total_data_product_sum.setter
    def cancel_total_data_product_sum(self, value: str):  # setter
        self.__cancel_total_data_product_sum = value

    @property
    def cancel_total_data_order_sum_amount(self):  # getter
        return self.__cancel_total_data_order_sum_amount

    @cancel_total_data_order_sum_amount.setter
    def cancel_total_data_order_sum_amount(self, value: str):  # setter
        self.__cancel_total_data_order_sum_amount = value

    @property
    def refund_total_data_product_sum(self):  # getter
        return self.__refund_total_data_product_sum

    @refund_total_data_product_sum.setter
    def refund_total_data_product_sum(self, value: str):  # setter
        self.__refund_total_data_product_sum = value

    @property
    def refund_total_data_order_sum_amount(self):  # getter
        return self.__refund_total_data_order_sum_amount

    @refund_total_data_order_sum_amount.setter
    def refund_total_data_order_sum_amount(self, value: str):  # setter
        self.__refund_total_data_order_sum_amount = value

    def to_print(self):
        print("상점이름", self.store_name)
