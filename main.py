import itertools
from string import digits, punctuation, ascii_letters
import win32com.client as client
from datetime import datetime
import time
from colorama import init, Fore


def brute_exel_doc():
    print("Hello people!")

    try:
        password_length = input("Введите длину пароля for example: 3-5:  ")
        password_length = [int(item) for item in password_length.split("-")]

        if len(password_length) != 2 or not all(isinstance(item, int) for item in password_length):
            raise ValueError("Неверный формат. Пожалуйста, введите длину пароля в формате 'число-число'.")
    except ValueError as e:
        print(f"Ошибка: {e}")
        print(brute_exel_doc())

    print("Если пароль содержит только цифры, введите: 1\nЕсли пароль содержит только буквы, введите: 2\n"
          "Если пароль содержит цифры и буквы введите: 3\nЕсли пароль содержит цифры, буквы и спецсимволы введите: 4")

    try:
        choice = int(input(": "))
        if choice == 1:
            possible_symbols = digits
        elif choice == 2:
            possible_symbols = ascii_letters
        elif choice == 3:
            possible_symbols = digits + ascii_letters
        elif choice == 4:
            possible_symbols = digits + ascii_letters + punctuation
        else:
            possible_symbols1 = "=========================\n" \
                                "Ты не правильно ввёл содержание пароля !\n" \
                                "=========================\n"
            print(f"{Fore.RED}{possible_symbols1}{Fore.RESET}")
            return (brute_exel_doc())
    except:
            possible_symbols1 = "=========================\n" \
                            "Ты не правильно ввёл длину пароля !\n" \
                            "=========================\n"
            print(f"{Fore.RED}{possible_symbols1}{Fore.RESET}")
            return (brute_exel_doc())


    # brute excel doc
    start_timestamp = time.time()
    print(f"Started at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")

    count = 0


    for pass_length in range(password_length[0], password_length[1] + 1):
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = "".join(password)
            count += 1
            #print(password)

            opened_doc = client.Dispatch("Excel.Application")

            try:
                opened_doc.Workbooks.Open(
                    # set password 1234
                    r"D:\даня\password_crack\test1.xlsx",
                    False,
                    True,
                    None,
                    password
                )
                time.sleep(0.1)
                print(f"Finished at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
                print(f"Password cracking time - {time.time() - start_timestamp}")

                password_correct = f"Attempt #{count} Password is: {password}"
                return(f"{Fore.GREEN}{password_correct}{Fore.RESET}")
            except:
                error = (f"Attempt #{count} Incorrect password: {password}")
                print(f"{Fore.RED}{error}{Fore.RESET}")
                pass


def main():
    print(brute_exel_doc())


if __name__ == '__main__':
    main()