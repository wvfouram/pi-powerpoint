#!/bin/bash
# Removes PowerPoint.
SRC="$(cd "$(dirname "$0")" && pwd)"
if [ "$(id -u)" = "0" ]; then
  echo "Run this as your normal user (not with sudo)."
  exit 1
fi

# put back the apps that opened these file types before
if [ -f "$HOME/.config/pi-powerpoint/old-defaults" ]; then
  while read -r t old; do
    [ -n "$old" ] && xdg-mime default "$old" "$t" 2>/dev/null || true
  done < "$HOME/.config/pi-powerpoint/old-defaults"
  rm -f "$HOME/.config/pi-powerpoint/old-defaults"
fi

sudo rm -rf /opt/pi-powerpoint /usr/local/bin/pi-powerpoint /usr/share/applications/pi-powerpoint.desktop \
  /usr/share/icons/hicolor/scalable/apps/pi-powerpoint.svg
sudo update-desktop-database -q /usr/share/applications 2>/dev/null || true
echo "PowerPoint removed. Your settings are still in ~/.config/pi-powerpoint (delete that folder to remove them)."
