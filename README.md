## Projeto hackathom - configurando desenvolvimento

Tenha python instalado <br>
Clone o repositório <br>
no terminal digite <br>

```bash
git checkout -b <branch-name> # para criar um branch novo com as suas mudanças
cd hackathom # para ir ate a pasta do repositorio
virtualenv env
source env/bin/activate
pip install -r requirements.txt
```
você deve ver um (env) no canto esquerdo do terminal, para desativar, digite deactivate.

```bash
(env) cd Hackathom #sim, de novo esse comando
(env) python3 manage.py runserver
```


