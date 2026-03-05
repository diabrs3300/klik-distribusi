"""
utils.py — Helper functions / utilities.
"""


def flash_errors(form):
    """Tampilkan semua error dari WTForms sebagai flash message."""
    from flask import flash
    for field, errors in form.errors.items():
        for error in errors:
            flash(f'{getattr(form, field).label.text}: {error}', 'danger')
