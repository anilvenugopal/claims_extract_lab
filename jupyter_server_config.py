# Project-local Jupyter config. Jupyter does NOT pick this up on its own, so
# pass it explicitly from the repo root:
#   jupyter lab --config=jupyter_server_config.py notebooks/...
# Local-only convenience: no token/password. Server stays bound to localhost.
c.IdentityProvider.token = ""
c.ServerApp.password = ""
c.ServerApp.ip = "127.0.0.1"
