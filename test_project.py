from project import name_check, level_check, col_check, contact_check, email_check


def test_name_check():
    assert name_check('', '') == (False, False)
    assert name_check('John', '') == (False, False)
    assert name_check('', 'Smith') == (False, False)
    assert name_check('john', 'smith') == ('John', 'Smith')


def test_level_check():
    assert level_check('') == False
    assert level_check('Frosh') == 'Frosh'


def test_col_check():
    assert col_check('Fine Arts') == False
    assert col_check('Science') == 'Science'
    assert col_check('Statistics') == 'Statistics'


def test_contact_check():
    assert contact_check('6309123456789') == '6309123456789'
    assert contact_check('630912345678901') == False
    assert contact_check('+6309123456789') == False


def test_email_check():
    assert email_check('david@harvard.edu') == 'david@harvard.edu'
    assert email_check('david.harvard.edu') == False
    assert email_check('david@harvard@edu') == False