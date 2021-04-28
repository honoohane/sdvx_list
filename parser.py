from functions import *
import json


def main():
    info_list = parse_16() + parse_17() + parse_18() + parse_19()

    with open('sdvx.json', 'w', encoding='utf-8') as js:
        json.dump(info_list, js, ensure_ascii=False)
    pd.DataFrame(info_list).to_csv("sdvx_list.csv", index=False, encoding='utf_8_sig', quoting=1)


if __name__ == "__main__":
    main()