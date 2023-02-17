from  interface import  *


def main():
    bank_lst = ['АО "АЛЬФА-БАНК"',
                'ПАО "БАНК УРАЛСИБ"']

    address_dict = dict(zip(bank_lst, ['119017, г.Москва, ул.Пятницкая, д. 40, стр.1',
                                       '119048, г. Москва, ул. Ефремова, 8']))

    form = Create_form(bank_lst=bank_lst, address_dict=address_dict)
    form.create_interface()

if __name__=='__main__':
    main()




