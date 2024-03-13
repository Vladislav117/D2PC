import os
import sys
import docx2pdf
import locale
import winreg


class Translations:
    __current_translation = dict()
    __all_translations = {
        "": dict()
    }

    @classmethod
    def add_translation(cls, language: str | tuple[str, ...] | list[str, ...], translation: dict) -> None:
        if isinstance(language, str):
            cls.__all_translations[language.lower()] = translation
            return
        for language_name in language:
            cls.__all_translations[language_name.lower()] = translation

    @classmethod
    def set_language(cls, language: str) -> None:
        if cls.__all_translations.__contains__(language.lower()):
            language = ""
        cls.__current_translation = cls.__all_translations.__getitem__(language)

    @classmethod
    def get_translation(cls, key: str) -> str:
        return cls.__current_translation.get(key, key)


Translations.add_translation(("", "en", "en_us"), {
    "docx_to_pdf_conversion_starts": "Conversion from DOCX to PDF starts...",
    "docx_to_pdf_conversion_ends": "Conversion from DOCX to PDF ends",
    "d2pc_is_being_installed": "D2PC is being installed...",
    "d2pc_has_already_been_installed": "D2PC has already been installed, deleting the entry from the register...",
    "d2pc_is_installed": "D2PC is installed",
})

Translations.add_translation(("ru", "ru_ru", "russian_russia"), {
    "docx_to_pdf_conversion_starts": "Конвертация из DOCX в PDF началась...",
    "docx_to_pdf_conversion_ends": "Конвертация из DOCX в PDF закончилась",
    "d2pc_is_being_installed": "D2PC устанавливается...",
    "d2pc_has_already_been_installed": "D2PC уже был установлен, удаление записи из регистра...",
    "d2pc_is_installed": "D2PC установлен",
})

Translations.set_language(locale.getlocale()[0].lower())

if sys.argv.__len__() > 1:
    docx_input = sys.argv[1]
    docx_output = docx_input.removesuffix(".docx") + ".pdf"
    print(Translations.get_translation("docx_to_pdf_conversion_starts"))
    docx2pdf.convert(docx_input, docx_output)
    print(Translations.get_translation("docx_to_pdf_conversion_ends"))
else:
    print(Translations.get_translation("d2pc_is_being_installed"))
    dot_docx_key = winreg.OpenKeyEx(winreg.HKEY_CLASSES_ROOT, r".docx")
    word_document_name = winreg.QueryValueEx(dot_docx_key, None)[0]
    word_document_key = winreg.OpenKeyEx(winreg.HKEY_CLASSES_ROOT, word_document_name)
    shell_key = winreg.OpenKeyEx(word_document_key, "shell")
    try:
        print(Translations.get_translation("d2pc_has_already_been_installed"))
        winreg.DeleteKey(shell_key, "D2PC")
    except:
        pass

    winreg.CreateKey(shell_key, "D2PC")
    d2pc_key = winreg.OpenKeyEx(shell_key, "D2PC", access=winreg.KEY_SET_VALUE)
    winreg.CreateKey(d2pc_key, "command")
    command_key = winreg.OpenKeyEx(d2pc_key, "command", access=winreg.KEY_SET_VALUE)

    winreg.SetValueEx(d2pc_key, "Position", 0, winreg.REG_SZ, "Bottom")
    winreg.SetValueEx(d2pc_key, "MUIVerb", 0, winreg.REG_SZ, "D2PC")
    winreg.SetValueEx(d2pc_key, "Icon", 0, winreg.REG_SZ, f"{os.getcwd()}\\D2PC.exe")
    winreg.SetValueEx(command_key, "", 0, winreg.REG_SZ, f"{os.getcwd()}\\D2PC.exe \"%1\"")
    print(Translations.get_translation("d2pc_is_installed"))
