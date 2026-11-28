#!/usr/bin/env bash
# Démarre le serveur FiveM.
# Utilisation : ./demarrer.sh            -> démarre directement le serveur
#               ./demarrer.sh txadmin    -> démarre le panneau web txAdmin (port 40120)
set -euo pipefail

cd "$(dirname "$0")"

if [ ! -f artifacts/run.sh ]; then
  echo "FXServer n'est pas installé : lance d'abord ./installer.sh"
  exit 1
fi

if [ "${1:-}" = "txadmin" ]; then
  echo "txAdmin démarre : ouvre http://adresse-du-serveur:40120 dans ton navigateur."
  echo "Choisis « Existing Server Data » et indique ce dossier : $(pwd)"
  exec bash artifacts/run.sh
fi

if [ ! -f secrets.cfg ]; then
  cp secrets.cfg.example secrets.cfg
fi
if grep -q 'colle_ta_cle_ici' secrets.cfg; then
  echo "Mets d'abord ta clé de licence dans secrets.cfg (voir README.md)."
  exit 1
fi

exec bash artifacts/run.sh +set onesync on +exec server.cfg
