#!/usr/bin/env bash
# Installe (ou met à jour) FXServer et les ressources de base de FiveM, pour Linux.
# Utilisation : ./installer.sh           -> version recommandée
#               ./installer.sh latest    -> toute dernière version
set -euo pipefail

cd "$(dirname "$0")"
CANAL="${1:-recommended}"
case "$CANAL" in
  recommended | latest) ;;
  *) echo "Utilisation : $0 [recommended|latest]"; exit 1 ;;
esac

for outil in curl tar xz; do
  if ! command -v "$outil" >/dev/null 2>&1; then
    echo "Il manque « $outil ». Sur Ubuntu / Debian : sudo apt install curl tar xz-utils"
    exit 1
  fi
done

TMP="$(pwd)/.installation"
rm -rf "$TMP"
mkdir -p "$TMP"
trap 'rm -rf "$TMP"' EXIT

echo "==> Recherche de la version « $CANAL » de FXServer..."
INFOS="$(curl -fsSL https://changelogs-live.fivem.net/api/changelog/versions/linux/server || true)"
INFOS="$(printf '%s' "$INFOS" | sed 's#\\/#/#g')"
URL="$(printf '%s' "$INFOS" | grep -o "\"${CANAL}_download\": *\"[^\"]*\"" | cut -d'"' -f4 || true)"
VERSION="$(printf '%s' "$INFOS" | grep -o "\"${CANAL}\": *\"\?[0-9]*" | grep -o '[0-9]*$' || true)"
if [ -z "$URL" ]; then
  echo "Impossible de trouver le lien de téléchargement."
  echo "Télécharge fx.tar.xz à la main sur https://runtime.fivem.net/artifacts/fivem/build_proot_linux/master/"
  echo "puis extrais-le dans le dossier « artifacts » (voir README.md, installation manuelle)."
  exit 1
fi

echo "==> Téléchargement de FXServer ${VERSION:-?}..."
curl -fL --progress-bar -o "$TMP/fx.tar.xz" "$URL"
rm -rf artifacts
mkdir -p artifacts
tar -xJf "$TMP/fx.tar.xz" -C artifacts

echo "==> Téléchargement des ressources de base (chat, spawnmanager...)..."
curl -fsSL -o "$TMP/cfx.tar.gz" https://github.com/citizenfx/cfx-server-data/archive/refs/heads/master.tar.gz
tar -xzf "$TMP/cfx.tar.gz" -C "$TMP"
rm -rf "resources/[cfx]"
mkdir -p "resources/[cfx]"
cp -R "$TMP"/cfx-server-data-*/resources/. "resources/[cfx]/"

if [ ! -f secrets.cfg ]; then
  cp secrets.cfg.example secrets.cfg
fi

echo
echo "Installation terminée !"
if grep -q 'colle_ta_cle_ici' secrets.cfg; then
  echo "Il reste à mettre ta clé de licence dans secrets.cfg (voir README.md),"
  echo "puis à lancer ./demarrer.sh"
else
  echo "Lance le serveur avec ./demarrer.sh"
fi
