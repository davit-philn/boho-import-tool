import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'Tools'))

from services.thdm_coverage import compute_coverage, muc_key, period_message

LINES = [
    dict(muc='2', item_id=2, code='B', name='Bb', qty=1, unit='HỆ'),
    dict(muc='1.10', item_id=11, code='A10', name='A10', qty=1, unit='HỆ'),
    dict(muc='1.2', item_id=12, code='A2', name='A2', qty=1, unit='HỆ'),
    dict(muc='3', item_id=None, code='', name='không mã', qty=1, unit='CÁI'),
]


def _bom(i, item, ver='1', active=True, code='X'):
    return dict(id=i, code=code, version=ver, item_id0=item, is_active=active)


def test_muc_key_sorts_numerically():
    ms = ['1.10', '1.2', '10', '2']
    assert sorted(ms, key=muc_key) == ['1.2', '1.10', '2', '10']


def test_coverage_marks_missing_and_sorts_by_muc():
    rows, s = compute_coverage(LINES, [_bom(1, 2), _bom(2, 12)])
    assert [r['muc'] for r in rows] == ['1.2', '1.10', '2', '3']
    assert [r['status'] for r in rows] == ['ok', 'missing', 'ok', 'missing']
    assert s == dict(total=4, ok=2, missing=2)


def test_inactive_bom_counts_as_missing():
    rows, s = compute_coverage(LINES[:1], [_bom(1, 2, active=False)])
    assert rows[0]['status'] == 'missing' and s['missing'] == 1


def test_highest_version_wins_per_product():
    rows, _ = compute_coverage(LINES[:1], [_bom(1, 2, ver='1', code='V1'), _bom(2, 2, ver='3', code='V3'),
                                           _bom(3, 2, ver='2', code='V2')])
    assert rows[0]['bom_id'] == 2 and rows[0]['bom_version'] == '3'


def test_period_message_levels():
    rows, s = compute_coverage(LINES[:1], [_bom(1, 2)])
    assert period_message('Đợt 1', s, rows)[1] == 'ok'
    rows, s = compute_coverage(LINES, [])
    text, level = period_message('Đợt 4', s, rows)
    assert level == 'error' and 'Mục 1.2, 1.10, 2, 3' in text
    assert period_message('Đợt 9', dict(total=0, ok=0, missing=0), [])[1] == 'empty'
