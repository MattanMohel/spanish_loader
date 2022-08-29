from bottle import template, get, route, post, request, run 
from src import loader

@post('/view')
def view():
    date = str(request.forms.get('date'))
    date, info = loader.get_date_info(date)
    return template('tpl/view', date=date, info=info)
    
@route('/')
def index():
    return template('tpl/index')

if __name__ == '__main__':
    run()