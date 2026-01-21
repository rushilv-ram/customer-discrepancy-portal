import requests

s = requests.Session()
# default credentials in environment are ADMIN_USER=admin and ADMIN_PASS=password
login = s.post('http://127.0.0.1:5000/login', data={'username': 'admin', 'password': 'password'})
print('Login:', login.status_code, login.url)

r = s.get('http://127.0.0.1:5000/api/stats')
print('/api/stats:', r.status_code)
print('Body sample:', r.text[:400])

r2 = s.get('http://127.0.0.1:5000/dashboard')
print('/dashboard:', r2.status_code)
open('tests/dashboard_page.html', 'wb').write(r2.content)
print('Saved dashboard_page.html')