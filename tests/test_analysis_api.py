from app import compute_analysis

sample = [
    {'Short':'3','Excess':'0','Damage':'','Mismatch Parts':'','Transporter Name':'T1','Customer Name':'C1','Warehouse':'W1','Invoice No':'I1','Submitted At':'2026-01-01'},
    {'Short':'2','Excess':'1','Damage':'X','Mismatch Parts':'','Transporter Name':'T2','Customer Name':'C1','Warehouse':'W1','Invoice No':'I2','Submitted At':'2026-01-02'},
    {'Short':'','Excess':'','Damage':'2','Mismatch Parts':'1','Transporter Name':'T1','Customer Name':'C2','Warehouse':'W2','Invoice No':'I3','Submitted At':'2026-01-03'},
]


def test_compute_short_by_transporter():
    res = compute_analysis(sample, metric='short', group_by='transporter')
    # Expect T1=3, T2=2
    assert any(r['group']=='T1' and r['value']==3 for r in res)
    assert any(r['group']=='T2' and r['value']==2 for r in res)


def test_compute_damage_by_customer():
    res = compute_analysis(sample, metric='damage', group_by='customer')
    # C1 had non-numeric Damage '' => 0; C2 had Damage=2
    assert any(r['group']=='C2' and r['value']==2 for r in res)


def test_date_filtering():
    res = compute_analysis(sample, metric='short', group_by='transporter', date_from='2026-01-02')
    # Only rows from 2026-01-02 and later: T2=2
    assert any(r['group']=='T2' and r['value']==2 for r in res)


def test_warehouse_filter():
    res = compute_analysis(sample, metric='short', group_by='transporter', warehouse_filter='W2')
    # Only W2 has a short 0 but damage 0; expect T1 not present
    assert all(r['group']!='T1' for r in res)
