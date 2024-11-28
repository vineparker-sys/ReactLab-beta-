import json
import os

class Localizer:
    def __init__(self, default_language="pt"):
        self.default_language = default_language
        self.translations = {}
        self.load_translations()

    def load_translations(self):
        """Loads all translations from the 'locales' folder."""
        base_path = os.path.join(os.path.dirname(__file__), "locales")
        for file in os.listdir(base_path):
            if file.endswith(".json"):
                lang_code = file.split(".")[0]
                with open(os.path.join(base_path, file), "r", encoding="utf-8") as f:
                    self.translations[lang_code] = json.load(f)

    def translate(self, key):
        """Translates a key based on the current language."""
        return self.translations.get(self.default_language, {}).get(key, key)

    def set_language(self, language):
        """Sets the current language."""
        if language in self.translations:
            self.default_language = language
        else:
            raise ValueError(f"Language '{language}' not supported.")
