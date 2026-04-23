import xlwings as xw

# Задание русского алфавита с буквой Ё
ALPHABET = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ"


def vigenere_encrypt(text: str, key: str) -> str:
    """Шифрует текст шифром Виженера по русскому алфавиту."""
    if not key:
        return text

    result = []
    key = key.upper()
    key_index = 0

    for ch in text:
        upper_ch = ch.upper()

        # Шифруем только русские буквы, остальные символы оставляем без изменений
        if upper_ch in ALPHABET:
            text_pos = ALPHABET.index(upper_ch)
            key_pos = ALPHABET.index(key[key_index % len(key)])
            cipher_char = ALPHABET[(text_pos + key_pos) % len(ALPHABET)]

            # Сохраняем исходный регистр буквы
            if ch.islower():
                cipher_char = cipher_char.lower()

            result.append(cipher_char)
            key_index += 1
        else:
            result.append(ch)

    return "".join(result)


def main():
    # Функция main() запускается кнопкой Run из add-in xlwings в Excel
    wb = xw.Book.caller()
    ws = wb.sheets["Шифрование"]

    # Читаем исходный текст и ключ из ячеек Excel
    source_text = ws.range("B3").value or ""
    key = ws.range("B5").value or ""

    # Выполняем шифрование
    encrypted_text = vigenere_encrypt(source_text, key)

    # Выводим результат обратно в книгу Excel
    ws.range("B4").value = encrypted_text

    # Дополнительно подсвечиваем ячейку с результатом после успешного шифрования
    ws.range("B4").color = (255, 242, 204)
