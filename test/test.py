import sys
sys.path.append('../')

import json
import pytest

from xl_to_json import convert_nested_json, gen_json
import pandas as pd


def test_nested_json():
    """
    Проверяем, что JSON строка преобразуется в словарь,
    а всё остальное остается как есть...
    """
    sheets = pd.read_excel('json.xlsx', sheet_name=None)
    sheet_1 = sheets['Лист1']
    sheet_2 = sheets['Лист2']

    convert_nested_json(sheet_1)
    assert isinstance(sheet_1['json'][0], dict)
    assert isinstance(sheet_1['json'][1], dict)

    convert_nested_json(sheet_2)
    assert isinstance(sheet_2['not json'][0], int)
    assert isinstance(sheet_2['not json'][1], float)
    assert isinstance(sheet_2['not json'][2], str)
    assert isinstance(sheet_2['not json'][3], str)


def test_gen_json():
    expected_json = {
        "Лист1": [
            {
                "json": {
                    "a": 123,
                    "b": {
                        "c": 333
                    }
                }
            },
            {
                "json": {
                    "name": "John",
                    "age": 31,
                    "city": "New York"
                }
            }
        ],
        "Лист2": [
            {
                "not json": 3333
            },
            {
                "not json": 123.123
            },
            {
                "not json": "просто строка"
            },
            {
                "not json": "{123}"
            }
        ]
    }

    sheets = pd.read_excel('json.xlsx', sheet_name=None)
    json_str = gen_json(sheets)
    assert json.loads(json_str) == expected_json


def test_time_format():
    sheets = pd.read_excel('time.xlsx', None)
    data = json.loads(gen_json(sheets))['Лист1']

    assert data[0]['time'] == '2015-10-20T00:00:00.000Z'
    assert data[1]['time'] == '2015-10-20T15:43:00.000Z'
    assert data[2]['time'] == '2011-11-11T11:11:11.000Z'
