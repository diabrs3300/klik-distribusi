"""
Unit tests untuk routes.
Jalankan dengan: pytest tests/
"""
import pytest
from app import create_app, db


@pytest.fixture
def app():
    """Buat instance aplikasi untuk testing."""
    app = create_app('testing')
    with app.app_context():
        db.create_all()
        yield app
        db.drop_all()


@pytest.fixture
def client(app):
    """Test client Flask."""
    return app.test_client()


def test_index_page(client):
    """Halaman beranda harus mengembalikan 200."""
    response = client.get('/')
    assert response.status_code == 200


def test_login_page(client):
    """Halaman login harus mengembalikan 200."""
    response = client.get('/auth/login')
    assert response.status_code == 200


def test_register_page(client):
    """Halaman register harus mengembalikan 200."""
    response = client.get('/auth/register')
    assert response.status_code == 200


def test_404_page(client):
    """URL yang tidak ada harus mengembalikan 404."""
    response = client.get('/halaman-tidak-ada')
    assert response.status_code == 404
