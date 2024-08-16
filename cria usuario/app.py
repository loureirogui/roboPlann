from flask import Flask, render_template, request, redirect, url_for
from criaUsuarios import criaUsuario
from atualizaObrigacao import atualizaObrigacao
from atualizaRegime import atualizaRegime
from cadastraEmpresa import cadastraEmpresa

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/start', methods=['POST'])
def start_automation():
    emailLogin = request.form['email']
    senhaLogin = request.form['password']
    file = request.files['file']
    file.save('uploaded_empresas.xlsx')
    # criaUsuario(emailLogin, senhaLogin)
    # atualizaObrigacao(emailLogin, senhaLogin)
    # atualizaRegime(emailLogin, senhaLogin)
    cadastraEmpresa(emailLogin, senhaLogin)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
    
    
