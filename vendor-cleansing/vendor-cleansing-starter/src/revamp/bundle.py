"""Compact internal transport: repeated PO text is stored once, never truncated."""
import json
import sys


def write_bundle(path, payload):
    strings = []
    indexes = {}

    def compact(value):
        if isinstance(value, str) and len(value) > 64:
            if value not in indexes:
                indexes[value] = len(strings)
                strings.append(value)
            return {'$text': indexes[value]}
        if isinstance(value, dict):
            return {key: compact(item) for key, item in value.items()}
        if isinstance(value, list):
            return [compact(item) for item in value]
        return value

    packed = compact(payload)
    packed['_text_pool'] = strings
    with path.open('w', encoding='utf-8') as handle:
        json.dump(packed, handle, ensure_ascii=False, separators=(',', ':'))


def read_bundle(path):
    with path.open(encoding='utf-8') as handle:
        packed = json.load(handle)
    strings = packed.pop('_text_pool', [])

    def expand(value):
        if isinstance(value, dict):
            if set(value) == {'$text'}:
                return strings[value['$text']]
            for key in value:
                value[key] = expand(value[key])
        elif isinstance(value, list):
            for index in range(len(value)):
                value[index] = expand(value[index])
        elif isinstance(value, str):
            return sys.intern(value)
        return value

    return expand(packed)
