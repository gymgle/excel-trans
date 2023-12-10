import random

import requests
from easygoogletranslate import EasyGoogleTranslate

import utils

google_lang = {
    'zh': 'zh-CN',
    'en': 'en'
}


class Translator:
    def __init__(self, from_lang, to_lang):
        self.from_lang = from_lang
        self.to_lang = to_lang

    def translate(self, query):
        raise NotImplementedError('translate method must be implemented in the child class')


class GoogleTranslate(Translator):
    def __init__(self, from_lang, to_lang):
        super().__init__(from_lang, to_lang)
        self.client = EasyGoogleTranslate(
            source_language=google_lang[from_lang],
            target_language=google_lang[to_lang],
            timeout=10
        )

    def translate(self, query):
        return self.client.translate(query)


class BaiduTranslate(Translator):
    def __init__(self, from_lang, to_lang):
        super().__init__(from_lang, to_lang)
        self.src = from_lang
        self.dst = to_lang
        self.app_id = utils.get_configuration('baidu', 'appid')
        self.app_key = utils.get_configuration('baidu', 'appkey')
        self.url = 'https://fanyi-api.baidu.com/api/trans/vip/translate'

    def translate(self, query):
        salt = random.randint(32768, 65536)
        sign = utils.md5_sign(self.app_id + query + str(salt) + self.app_key)
        payload = {'appid': self.app_id, 'q': query, 'from': self.src, 'to': self.dst, 'salt': salt, 'sign': sign}
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        rsp = requests.post(self.url, params=payload, headers=headers)
        result = rsp.json().get('trans_result', [])
        if not result:
            return ''
        result_list = []
        for i in result:
            result_list.append(i.get('dst', ''))
        result_text = '\n'.join(result_list)
        return result_text


def create_translator(translator_type, from_lang, to_lang):
    if translator_type == 'google':
        return GoogleTranslate(from_lang, to_lang)
    elif translator_type == 'baidu':
        return BaiduTranslate(from_lang, to_lang)
    else:
        raise ValueError("Invalid translator type")


if __name__ == '__main__':
    test_content_zh = '这是一个测试，非常简单的测试。\n日期：2023.12.25'
    test_content_en = 'This is a test, a very simple test.\nDate: 2023.12.25'
    client = create_translator('google', 'zh', 'en')
    print(f'google zh->en: {client.translate(test_content_zh)}')

    client = create_translator('google', 'en', 'zh')
    print(f'google en->zh: {client.translate(test_content_en)}')

    client = create_translator('baidu', 'zh', 'en')
    print(f'baidu zh->en: {client.translate(test_content_zh)}')

    client = create_translator('baidu', 'en', 'zh')
    print(f'baidu en->zh: {client.translate(test_content_en)}')
