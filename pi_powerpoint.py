#!/usr/bin/python3
# PowerPoint for Raspberry Pi: a Windows 11 style presentation app that opens and saves .pptx files.
import os
import sys
import re
import json
import math
import time
import shutil
import signal
import threading
import subprocess

import gi
gi.require_version("Gtk", "3.0")
gi.require_version("Gdk", "3.0")
gi.require_version("GdkPixbuf", "2.0")
from gi.repository import Gtk, Gdk, GLib, GdkPixbuf, Pango, Gio, GObject  # noqa: E402,F401
import cairo  # noqa: E402

HOME = os.path.expanduser("~")
C_ACCENT = "#4CC2FF"
def hexc(h, a=1.0):
    h = h.lstrip("#")
    return (int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255, a)


def _rr(cr, x, y, w, h, r):
    r = min(r, w / 2, h / 2)
    cr.new_sub_path()
    cr.arc(x + w - r, y + r, r, -math.pi / 2, 0)
    cr.arc(x + w - r, y + h - r, r, 0, math.pi / 2)
    cr.arc(x + r, y + h - r, r, math.pi / 2, math.pi)
    cr.arc(x + r, y + r, r, math.pi, 1.5 * math.pi)
    cr.close_path()


def _line(cr, *pts):
    cr.move_to(*pts[0])
    for p in pts[1:]:
        cr.line_to(*p)
    cr.stroke()


def draw_glyph(cr, name, s, col):
    """Monochrome line icons for buttons."""
    cr.save()
    cr.scale(s / 16.0, s / 16.0)
    cr.set_source_rgba(*col)
    cr.set_line_width(1.15)
    cr.set_line_cap(cairo.LINE_CAP_ROUND)
    cr.set_line_join(cairo.LINE_JOIN_ROUND)
    pi = math.pi
    L = lambda *p: _line(cr, *p)  # noqa: E731
    if name == "back":
        L((13.5, 8), (2.5, 8))
        L((7, 3.5), (2.5, 8), (7, 12.5))
    elif name == "forward":
        L((2.5, 8), (13.5, 8))
        L((9, 3.5), (13.5, 8), (9, 12.5))
    elif name == "up":
        L((8, 13.5), (8, 2.5))
        L((3.5, 7), (8, 2.5), (12.5, 7))
    elif name == "refresh":
        cr.arc(8, 8, 5.5, -pi * 0.35, pi * 1.55)
        cr.stroke()
        L((12.9, 1.8), (12.6, 4.9), (9.6, 4.6))
    elif name == "plus":
        L((8, 2.5), (8, 13.5))
        L((2.5, 8), (13.5, 8))
    elif name == "new":
        cr.set_source_rgba(*hexc(C_ACCENT))
        cr.arc(8, 8, 6.5, 0, 2 * pi)
        cr.stroke()
        L((8, 4.8), (8, 11.2))
        L((4.8, 8), (11.2, 8))
    elif name == "cut":
        cr.arc(4.5, 12, 2, 0, 2 * pi)
        cr.stroke()
        cr.arc(11.5, 12, 2, 0, 2 * pi)
        cr.stroke()
        L((5.8, 10.5), (11, 1.8))
        L((10.2, 10.5), (5, 1.8))
    elif name == "copy":
        _rr(cr, 5, 1.5, 9, 10.5, 1.5)
        cr.stroke()
        cr.move_to(3, 4.5)
        cr.line_to(3, 13)
        cr.arc_negative(4.5, 13, 1.5, pi, pi / 2)
        cr.line_to(11, 14.5)
        cr.stroke()
    elif name == "paste":
        _rr(cr, 2.5, 3, 11, 11.5, 1.5)
        cr.stroke()
        _rr(cr, 5.5, 1.5, 5, 3, 1)
        cr.stroke()
        L((5.5, 8), (10.5, 8))
        L((5.5, 11), (9, 11))
    elif name == "rename":
        _rr(cr, 1.5, 4.5, 13, 7, 1.3)
        cr.stroke()
        L((4, 8), (8, 8))
        L((10.5, 2.5), (10.5, 13.5))
        L((9, 2.5), (12, 2.5))
        L((9, 13.5), (12, 13.5))
    elif name == "delete" or name == "trash":
        L((2.5, 4), (13.5, 4))
        L((6, 4), (6.5, 2), (9.5, 2), (10, 4))
        cr.move_to(4, 4)
        cr.line_to(4.8, 14)
        cr.line_to(11.2, 14)
        cr.line_to(12, 4)
        cr.stroke()
        L((6.8, 6.8), (7, 11.5))
        L((9.2, 6.8), (9, 11.5))
    elif name == "sort":
        L((4.5, 2.5), (4.5, 13.5))
        L((2, 5), (4.5, 2.5), (7, 5))
        L((11.5, 2.5), (11.5, 13.5))
        L((9, 11), (11.5, 13.5), (14, 11))
    elif name == "view":
        for y in (3, 8, 13):
            L((2, y), (14, y))
    elif name == "more":
        for x in (3, 8, 13):
            cr.arc(x, 8, 1.1, 0, 2 * pi)
            cr.fill()
    elif name == "chevdown":
        L((4, 6), (8, 10), (12, 6))
    elif name == "chevright":
        L((6, 4), (10, 8), (6, 12))
    elif name == "close":
        L((3.5, 3.5), (12.5, 12.5))
        L((12.5, 3.5), (3.5, 12.5))
    elif name == "min":
        L((3, 8), (13, 8))
    elif name == "max":
        _rr(cr, 3, 3, 10, 10, 1.2)
        cr.stroke()
    elif name == "restore":
        _rr(cr, 3, 5, 8, 8, 1.2)
        cr.stroke()
        L((5, 3), (12, 3), (13, 4), (13, 11))
    elif name == "search":
        cr.arc(6.6, 6.6, 4.6, 0, 2 * pi)
        cr.stroke()
        L((10, 10), (14, 14))
    elif name == "pane":
        _rr(cr, 1.5, 2.5, 13, 11, 1.5)
        cr.stroke()
        L((10, 2.5), (10, 13.5))
    elif name == "terminal":
        _rr(cr, 1.5, 2.5, 13, 11, 1.5)
        cr.stroke()
        L((4, 6), (6.5, 8), (4, 10))
        L((8, 10.5), (11.5, 10.5))
    elif name == "properties":
        _rr(cr, 2.5, 1.5, 11, 13, 1.5)
        cr.stroke()
        for y in (5, 8, 11):
            cr.arc(5.2, y, 0.8, 0, 2 * pi)
            cr.fill()
            L((7.2, y), (11, y))
    elif name == "pin":
        cr.move_to(6, 2)
        cr.line_to(10, 2)
        cr.line_to(9.5, 6.5)
        cr.line_to(12, 9)
        cr.line_to(4, 9)
        cr.line_to(6.5, 6.5)
        cr.close_path()
        cr.stroke()
        L((8, 9), (8, 14))
    elif name == "pin-fill":
        cr.set_line_width(1.0)
        cr.translate(8, 8)
        cr.rotate(pi / 4)
        cr.translate(-8, -8)
        cr.move_to(6, 2)
        cr.line_to(10, 2)
        cr.line_to(9.5, 6.5)
        cr.line_to(12, 9)
        cr.line_to(4, 9)
        cr.line_to(6.5, 6.5)
        cr.close_path()
        cr.stroke()
        L((8, 9), (8, 14))
    elif name == "folder":
        cr.move_to(1.5, 4)
        cr.line_to(6, 4)
        cr.line_to(7.5, 5.5)
        cr.line_to(14.5, 5.5)
        cr.line_to(14.5, 13)
        cr.line_to(1.5, 13)
        cr.close_path()
        cr.stroke()
    elif name == "file":
        cr.move_to(3.5, 1.5)
        cr.line_to(9.5, 1.5)
        cr.line_to(12.5, 4.5)
        cr.line_to(12.5, 14.5)
        cr.line_to(3.5, 14.5)
        cr.close_path()
        cr.stroke()
        L((9.5, 1.5), (9.5, 4.5), (12.5, 4.5))
    elif name == "zip":
        _rr(cr, 2.5, 1.5, 11, 13, 1.5)
        cr.stroke()
        for y in (3, 5, 7):
            L((7, y), (9, y))
        _rr(cr, 6.5, 9, 3, 3.5, 0.6)
        cr.stroke()
    elif name == "extract":
        _rr(cr, 2.5, 6.5, 11, 8, 1.5)
        cr.stroke()
        L((8, 1.5), (8, 10))
        L((5.5, 7.5), (8, 10), (10.5, 7.5))
    elif name == "restoreitem":
        cr.arc(8.5, 8.5, 5, -pi * 0.9, pi * 0.6)
        cr.stroke()
        L((2, 3.5), (3.6, 6.6), (6.6, 5.4))
    elif name == "open":
        _rr(cr, 1.5, 1.5, 13, 13, 2)
        cr.stroke()
        L((7, 9), (12.5, 3.5))
        L((8.5, 3.5), (12.5, 3.5), (12.5, 7.5))
    elif name == "newtab":
        _rr(cr, 1.5, 3.5, 13, 10, 1.5)
        cr.stroke()
        L((1.5, 6), (8, 6))
        L((8, 3.5), (8, 6))
    elif name == "check":
        L((3, 8.5), (6.5, 12), (13, 4.5))
    elif name == "eject":
        cr.move_to(8, 3)
        cr.line_to(13, 9)
        cr.line_to(3, 9)
        cr.close_path()
        cr.stroke()
        L((3, 12.5), (13, 12.5))
    elif name == "details-view":
        for y in (3.5, 8, 12.5):
            _rr(cr, 1.5, y - 1.2, 2.4, 2.4, 0.5)
            cr.stroke()
            L((6, y), (14.5, y))
    elif name == "icons-view":
        for x in (2, 9):
            for y in (2, 9):
                _rr(cr, x, y, 5, 5, 1.2)
                cr.stroke()
    elif name == "selectall":
        _rr(cr, 2, 2, 12, 12, 1.5)
        cr.stroke()
        L((5, 8), (7.2, 10.2), (11, 6))
    elif name == "info":
        cr.arc(8, 8, 6, 0, 2 * pi)
        cr.stroke()
        L((8, 7), (8, 11.5))
        cr.arc(8, 4.8, 0.9, 0, 2 * pi)
        cr.fill()
    elif name == "share":
        cr.arc(4, 8, 1.8, 0, 2 * pi)
        cr.stroke()
        cr.arc(12, 3.8, 1.8, 0, 2 * pi)
        cr.stroke()
        cr.arc(12, 12.2, 1.8, 0, 2 * pi)
        cr.stroke()
        L((5.6, 7.1), (10.4, 4.7))
        L((5.6, 8.9), (10.4, 11.3))
    elif name == "network":
        _rr(cr, 5, 1.5, 6, 4.5, 1)
        cr.stroke()
        L((8, 6), (8, 9))
        L((3, 9), (13, 9))
        L((3, 9), (3, 11))
        L((13, 9), (13, 11))
        _rr(cr, 1, 11, 4, 3.5, 0.8)
        cr.stroke()
        _rr(cr, 11, 11, 4, 3.5, 0.8)
        cr.stroke()
    elif name == "link":
        cr.arc(5.5, 10.5, 3, pi * 0.25, pi * 1.75)
        cr.stroke()
        cr.arc(10.5, 5.5, 3, pi * 1.25, pi * 0.75)
        cr.stroke()
        L((6.5, 9.5), (9.5, 6.5))
    else:
        draw_glyph_extra(cr, name, L, pi)
    cr.restore()




# ----------------------------------------------------------------------------
# Shared Windows 11 dark style
# ----------------------------------------------------------------------------
BASE_CSS = """
* { font-family: "Segoe UI Variable Text", "Segoe UI", "Selawik", "Open Sans", "Noto Sans", sans-serif; }
window.fx, window.fx-dialog { background-color: #202020; color: #ffffff; font-size: 12px; }
.fx label { color: #ffffff; }
headerbar.fx-title { background: #202020; border: none; box-shadow: none; min-height: 40px; padding: 0; margin: 0; }
headerbar.fx-title:backdrop { background: #202020; }
.app-title { font-size: 12px; color: #ffffff; margin-left: 4px; }
.app-title:backdrop { color: rgba(255,255,255,0.55); }
button.wc { background: transparent; border: none; border-radius: 0; min-width: 46px; min-height: 32px; padding: 0; box-shadow: none; }
button.wc:hover { background-color: rgba(255,255,255,0.06); }
button.wc.close:hover { background-color: #C42B1C; }
.tabstrip { margin-top: 6px; margin-left: 6px; }
button.tab { outline-color: transparent; background: transparent; border: none; border-radius: 8px 8px 0 0; min-height: 34px;
   padding: 0 6px 0 12px; box-shadow: none; margin-right: 1px; }
button.tab:hover { background-color: rgba(255,255,255,0.04); }
button.tab.active { background-color: #2C2C2C; }
button.tab label { font-size: 12px; }
button.tab-x { background: transparent; border: none; border-radius: 4px; min-width: 24px; min-height: 24px; padding: 0; box-shadow: none; }
button.tab-x:hover { background-color: rgba(255,255,255,0.08); }
button.tab-new { background: transparent; border: none; border-radius: 4px; min-width: 32px; min-height: 28px; padding: 0;
   box-shadow: none; margin-top: 4px; margin-left: 2px; }
button.tab-new:hover { background-color: rgba(255,255,255,0.06); }
.bar { background-color: #2C2C2C; padding: 4px 8px; }
.bar-bottom { border-bottom: 1px solid rgba(255,255,255,0.06); }
button.tb, button.nb { background: transparent; border: none; border-radius: 4px; min-height: 32px; padding: 0 9px; box-shadow: none; color: #ffffff; }
button.nb { min-width: 34px; padding: 0; }
button.tb:hover, button.nb:hover { background-color: rgba(255,255,255,0.06); }
button.tb:active, button.nb:active { background-color: rgba(255,255,255,0.04); }
button.tb:disabled label { color: rgba(255,255,255,0.36); }
button.tb label { font-size: 12px; }
button.tb.on, button.nb.on { background-color: rgba(255,255,255,0.09); }
.vsep { background-color: rgba(255,255,255,0.08); min-width: 1px; margin: 6px 6px; }
.status { background-color: #202020; padding: 4px 12px; border-top: 1px solid rgba(255,255,255,0.06); }
.status label { font-size: 12px; color: rgba(255,255,255,0.786); }
.sec { color: rgba(255,255,255,0.786); }
.ter { color: rgba(255,255,255,0.545); }
.big-title { font-size: 20px; font-weight: 600; }
.sect-title { font-size: 14px; font-weight: 600; }
entry.field, entry.fx-search { background-color: rgba(255,255,255,0.06); color: #ffffff; border: 1px solid rgba(255,255,255,0.07);
   border-bottom: 1px solid rgba(255,255,255,0.54); border-radius: 4px; min-height: 30px; padding: 0 10px; box-shadow: none;
   font-size: 12px; caret-color: #ffffff; }
entry.field:focus, entry.fx-search:focus { background-color: #1c1c1c; border-bottom: 2px solid #4CC2FF; }
entry selection, textview text selection, entry selection:focus, textview text selection:focus { background-color: #0078D4; color: #ffffff; }
scrollbar { background: transparent; border: none; }
scrollbar slider { background-color: rgba(255,255,255,0.45); border-radius: 6px; min-width: 3px; min-height: 3px; border: none; }
scrollbar:hover slider { min-width: 6px; min-height: 6px; background-color: rgba(255,255,255,0.55); }
scrollbar trough { background: transparent; border: none; }
scale trough { background-color: rgba(255,255,255,0.3); border: none; border-radius: 2px; min-height: 4px; min-width: 4px; }
scale highlight { background-color: #4CC2FF; border: none; border-radius: 2px; min-height: 4px; min-width: 4px; }
scale slider { background-color: #4CC2FF; border: 5px solid #454545; border-radius: 50%; min-width: 12px; min-height: 12px; box-shadow: none; }
scale slider:hover { border-width: 4px; }
menu, .menu, .context-menu { background-color: #2C2C2C; color: #ffffff; border: 1px solid rgba(0,0,0,0.45); border-radius: 8px; padding: 4px; }
menu menuitem { padding: 6px 12px; border-radius: 4px; color: #ffffff; min-height: 18px; }
menu menuitem:hover { background-color: rgba(255,255,255,0.06); }
menu menuitem:disabled label { color: rgba(255,255,255,0.36); }
menu menuitem label { color: #ffffff; font-size: 12px; }
menu menuitem arrow { color: rgba(255,255,255,0.786); }
menu separator { background-color: rgba(255,255,255,0.08); margin: 4px 0; min-height: 1px; }
window.fx button, window.fx-dialog button { background-image: none; text-shadow: none; -gtk-icon-shadow: none; }
button.primary { background-color: #4CC2FF; color: #000000; border: none; border-radius: 4px; min-height: 30px; padding: 0 16px; box-shadow: none; }
button.primary label { color: #000000; }
button.primary:hover { background-color: #48B1E8; }
button.secondary { background-color: rgba(255,255,255,0.06); color: #ffffff; border: 1px solid rgba(255,255,255,0.09); border-radius: 4px; min-height: 30px; padding: 0 16px; box-shadow: none; }
button.secondary:hover { background-color: rgba(255,255,255,0.09); }
.dlg-body { background-color: #2B2B2B; padding: 24px; }
.dlg-foot { background-color: #202020; padding: 22px 24px; border-top: 1px solid rgba(0,0,0,0.3); }
.dlg-title { font-size: 20px; font-weight: 600; }
.fx-dialog label { color: #ffffff; }
.fx-dialog label.sec { color: rgba(255,255,255,0.786); }
.fx-dialog label.ter { color: rgba(255,255,255,0.545); }
checkbutton check { background-color: rgba(0,0,0,0.1); border: 1px solid rgba(255,255,255,0.6); border-radius: 4px; min-width: 18px; min-height: 18px; -gtk-icon-source: none; color: #000000; }
checkbutton check:checked { background-color: #4CC2FF; border-color: #4CC2FF; -gtk-icon-source: -gtk-icontheme("object-select-symbolic"); }
checkbutton label { color: #ffffff; font-size: 12px; margin-left: 6px; }
radiobutton radio { background-color: rgba(0,0,0,0.1); border: 1px solid rgba(255,255,255,0.6); border-radius: 50%; min-width: 18px; min-height: 18px; -gtk-icon-source: none; }
radiobutton radio:checked { background-color: #000000; border: 5px solid #4CC2FF; }
radiobutton label { color: #ffffff; font-size: 12px; margin-left: 6px; }
switch { background-color: transparent; border: 1px solid rgba(255,255,255,0.6); border-radius: 12px; min-width: 40px; min-height: 18px; }
switch:checked { background-color: #4CC2FF; border-color: #4CC2FF; }
switch slider { background-color: rgba(255,255,255,0.8); border-radius: 50%; min-width: 12px; min-height: 12px; margin: 3px; border: none; box-shadow: none; }
switch:checked slider { background-color: #000000; }
switch image { color: transparent; }
combobox button, combobox box { background-color: rgba(255,255,255,0.06); color: #ffffff; border: 1px solid rgba(255,255,255,0.07); border-radius: 4px; box-shadow: none; min-height: 30px; }
combobox button cellview, combobox button label, combobox arrow { color: #ffffff; }
spinbutton { background-color: rgba(255,255,255,0.06); color: #ffffff; border: 1px solid rgba(255,255,255,0.07); border-radius: 4px; box-shadow: none; }
spinbutton entry { background: transparent; border: none; box-shadow: none; color: #ffffff; }
spinbutton button { background: transparent; border: none; box-shadow: none; color: #ffffff; }
tooltip, tooltip.background { background-color: #2C2C2C; border: 1px solid rgba(0,0,0,0.4); border-radius: 4px; }
tooltip label { color: #ffffff; font-size: 12px; }
.set-group { background-color: rgba(255,255,255,0.05); border-radius: 4px; border: 1px solid rgba(0,0,0,0.2); padding: 12px 16px; }
.set-head { font-size: 13px; }
.set-desc { font-size: 12px; color: rgba(255,255,255,0.6); }
"""


def apply_css(extra=""):
    prov = Gtk.CssProvider()
    prov.load_from_data((BASE_CSS + extra).encode())
    Gtk.StyleContext.add_provider_for_screen(Gdk.Screen.get_default(), prov,
                                             Gtk.STYLE_PROVIDER_PRIORITY_APPLICATION + 10)
    Gtk.Settings.get_default().set_property("gtk-application-prefer-dark-theme", True)


def draw_glyph_extra(cr, name, L, pi):
    if name == "save":
        _rr(cr, 2, 2, 12, 12, 1.5)
        cr.stroke()
        _rr(cr, 4.5, 2, 6, 4, 0.5)
        cr.stroke()
        _rr(cr, 4.5, 9, 7, 5, 0.5)
        cr.stroke()
    elif name == "undo":
        cr.move_to(3, 7)
        cr.curve_to(6, 3.5, 13, 4.5, 13, 10)
        cr.stroke()
        L((3, 3), (3, 7), (7, 7))
    elif name == "redo":
        cr.move_to(13, 7)
        cr.curve_to(10, 3.5, 3, 4.5, 3, 10)
        cr.stroke()
        L((13, 3), (13, 7), (9, 7))
    elif name == "replace":
        L((2, 5), (10, 5))
        L((7.5, 2.5), (10, 5), (7.5, 7.5))
        L((14, 11), (6, 11))
        L((8.5, 8.5), (6, 11), (8.5, 13.5))
    elif name == "zoomin":
        cr.arc(6.6, 6.6, 4.6, 0, 2 * pi)
        cr.stroke()
        L((10, 10), (14, 14))
        L((4.6, 6.6), (8.6, 6.6))
        L((6.6, 4.6), (6.6, 8.6))
    elif name == "zoomout":
        cr.arc(6.6, 6.6, 4.6, 0, 2 * pi)
        cr.stroke()
        L((10, 10), (14, 14))
        L((4.6, 6.6), (8.6, 6.6))
    elif name == "fit":
        L((2, 5.5), (2, 2), (5.5, 2))
        L((10.5, 2), (14, 2), (14, 5.5))
        L((14, 10.5), (14, 14), (10.5, 14))
        L((5.5, 14), (2, 14), (2, 10.5))
    elif name == "wrap":
        L((2, 3), (14, 3))
        cr.move_to(2, 8)
        cr.line_to(11.5, 8)
        cr.curve_to(15, 8, 15, 13, 11.5, 13)
        cr.line_to(8, 13)
        cr.stroke()
        L((9.5, 11.3), (8, 13), (9.5, 14.7))
        L((2, 13), (5, 13))
    elif name == "play":
        cr.move_to(4.5, 2.5)
        cr.line_to(13, 8)
        cr.line_to(4.5, 13.5)
        cr.close_path()
        cr.fill()
    elif name == "pause":
        _rr(cr, 3.5, 2.5, 3, 11, 0.8)
        cr.fill()
        _rr(cr, 9.5, 2.5, 3, 11, 0.8)
        cr.fill()
    elif name in ("next", "prev"):
        if name == "prev":
            cr.translate(16, 0)
            cr.scale(-1, 1)
        cr.move_to(3, 3)
        cr.line_to(10.5, 8)
        cr.line_to(3, 13)
        cr.close_path()
        cr.fill()
        _rr(cr, 11, 3, 2, 10, 0.6)
        cr.fill()
    elif name == "shuffle":
        L((2, 4), (5, 4), (11, 12), (14, 12))
        L((2, 12), (5, 12), (11, 4), (14, 4))
        L((12, 2), (14, 4), (12, 6))
        L((12, 10), (14, 12), (12, 14))
    elif name in ("repeat", "repeat1"):
        cr.move_to(3, 9)
        cr.line_to(3, 6)
        cr.arc(5, 6, 2, pi, 1.5 * pi)
        cr.line_to(13, 4)
        cr.stroke()
        L((11, 2), (13, 4), (11, 6))
        cr.move_to(13, 7)
        cr.line_to(13, 10)
        cr.arc(11, 10, 2, 0, 0.5 * pi)
        cr.line_to(3, 12)
        cr.stroke()
        L((5, 10), (3, 12), (5, 14))
        if name == "repeat1":
            cr.set_font_size(5)
            cr.move_to(6.8, 9.8)
            cr.show_text("1")
    elif name in ("volume", "mute"):
        cr.move_to(1.5, 6)
        cr.line_to(4.5, 6)
        cr.line_to(8, 2.5)
        cr.line_to(8, 13.5)
        cr.line_to(4.5, 10)
        cr.line_to(1.5, 10)
        cr.close_path()
        cr.stroke()
        if name == "mute":
            L((10.5, 6), (14.5, 10))
            L((14.5, 6), (10.5, 10))
        else:
            cr.arc(8, 8, 3.5, -pi / 3.5, pi / 3.5)
            cr.stroke()
            cr.arc(8, 8, 6.2, -pi / 3.5, pi / 3.5)
            cr.stroke()
    elif name == "fullscreen":
        L((2, 6), (2, 2), (6, 2))
        L((10, 2), (14, 2), (14, 6))
        L((14, 10), (14, 14), (10, 14))
        L((6, 14), (2, 14), (2, 10))
        L((2, 2), (6, 6))
        L((14, 14), (10, 10))
    elif name == "unfullscreen":
        L((6, 2), (6, 6), (2, 6))
        L((10, 2), (10, 6), (14, 6))
        L((10, 14), (10, 10), (14, 10))
        L((6, 14), (6, 10), (2, 10))
    elif name == "rotate":
        cr.arc(8, 9, 5, pi * 1.15, pi * 2.85)
        cr.stroke()
        L((2.5, 2.5), (3.4, 6.4), (7.2, 5.4))
    elif name == "heart":
        cr.move_to(8, 13.5)
        cr.curve_to(1, 8, 1.5, 2.5, 5, 2.5)
        cr.curve_to(6.5, 2.5, 7.5, 3.5, 8, 4.5)
        cr.curve_to(8.5, 3.5, 9.5, 2.5, 11, 2.5)
        cr.curve_to(14.5, 2.5, 15, 8, 8, 13.5)
        cr.stroke()
    elif name == "pen":
        L((2.5, 13.5), (3.2, 10.5), (11, 2.7), (13.3, 5), (5.5, 12.8), (2.5, 13.5))
        L((9.5, 4.2), (11.8, 6.5))
    elif name == "highlighter":
        L((3, 13.5), (13, 13.5))
        cr.move_to(5, 11)
        cr.line_to(11, 3)
        cr.line_to(13.5, 5)
        cr.line_to(7.5, 12.5)
        cr.line_to(5, 12.5)
        cr.close_path()
        cr.stroke()
    elif name == "eraser":
        cr.move_to(2.5, 10)
        cr.line_to(9, 3.5)
        cr.line_to(13.5, 8)
        cr.line_to(8, 13.5)
        cr.line_to(5.5, 13.5)
        cr.close_path()
        cr.stroke()
        L((6, 6.5), (10.5, 11))
        L((8, 13.5), (14, 13.5))
    elif name == "crop":
        L((4.5, 1.5), (4.5, 11.5), (14.5, 11.5))
        L((1.5, 4.5), (11.5, 4.5), (11.5, 14.5))
    elif name == "camera":
        _rr(cr, 1.5, 4, 13, 9.5, 1.6)
        cr.stroke()
        L((5, 4), (6, 2.5), (10, 2.5), (11, 4))
        cr.arc(8, 8.8, 2.6, 0, 2 * pi)
        cr.stroke()
    elif name == "timer":
        cr.arc(8, 9, 5.5, 0, 2 * pi)
        cr.stroke()
        L((8, 9), (8, 5.8))
        L((6.3, 1.8), (9.7, 1.8))
    elif name == "settings":
        n = 8
        cr.new_path()
        for i in range(n * 2):
            a = i * pi / n
            r = 6.6 if i % 2 == 0 else 5.0
            a1, a2 = a - pi / (n * 2.4), a + pi / (n * 2.4)
            cr.line_to(8 + r * math.cos(a1), 8 + r * math.sin(a1))
            cr.line_to(8 + r * math.cos(a2), 8 + r * math.sin(a2))
        cr.close_path()
        cr.stroke()
        cr.arc(8, 8, 2.2, 0, 2 * pi)
        cr.stroke()
    elif name == "history":
        cr.arc(8, 8, 6, pi * 1.05, pi * 2.85)
        cr.stroke()
        L((1.2, 6.2), (2.3, 8.6), (4.6, 7.3))
        L((8, 4.6), (8, 8.2), (10.5, 9.7))
    elif name == "image":
        _rr(cr, 1.5, 2.5, 13, 11, 1.5)
        cr.stroke()
        L((1.5, 11), (5.5, 7), (9, 10.5), (11, 8.5), (14.5, 12))
        cr.arc(10.5, 5.5, 1.2, 0, 2 * pi)
        cr.stroke()
    elif name == "music":
        cr.arc(4.5, 12, 2, 0, 2 * pi)
        cr.stroke()
        cr.arc(12, 10.5, 2, 0, 2 * pi)
        cr.stroke()
        L((6.5, 12), (6.5, 3.5), (14, 2), (14, 10.5))
    elif name == "video":
        _rr(cr, 1.5, 3.5, 9.5, 9, 1.5)
        cr.stroke()
        L((11, 6.5), (14.5, 4.5), (14.5, 11.5), (11, 9.5))
    elif name == "queue":
        for y in (3, 7):
            L((2, y), (14, y))
        L((2, 11), (8, 11))
        cr.move_to(10.5, 9)
        cr.line_to(14, 11.5)
        cr.line_to(10.5, 14)
        cr.close_path()
        cr.fill()
    elif name == "edit":
        L((2.5, 13.5), (3.2, 10.5), (11, 2.7), (13.3, 5), (5.5, 12.8), (2.5, 13.5))
    elif name == "slideshow":
        _rr(cr, 1.5, 2.5, 13, 9, 1.5)
        cr.stroke()
        cr.move_to(6.5, 5)
        cr.line_to(10, 7)
        cr.line_to(6.5, 9)
        cr.close_path()
        cr.fill()
        L((5, 14), (11, 14))
    elif name == "print":
        _rr(cr, 1.5, 6, 13, 6, 1.2)
        cr.stroke()
        L((4.5, 6), (4.5, 1.5), (11.5, 1.5), (11.5, 6))
        _rr(cr, 4.5, 9.5, 7, 5, 0.5)
        cr.stroke()
    elif name == "wallpaper":
        _rr(cr, 1.5, 2.5, 13, 9, 1.2)
        cr.stroke()
        L((1.5, 9.5), (5, 6), (8, 9), (10, 7), (14.5, 10.5))
        L((5.5, 14), (10.5, 14))
        L((8, 11.5), (8, 14))
    elif name == "list":
        for y in (3.5, 8, 12.5):
            cr.arc(2.5, y, 0.9, 0, 2 * pi)
            cr.fill()
            L((5.5, y), (14, y))
    elif name == "add":
        L((8, 2.5), (8, 13.5))
        L((2.5, 8), (13.5, 8))
    elif name == "split":
        _rr(cr, 1.5, 2.5, 13, 11, 1.5)
        cr.stroke()
        L((8, 2.5), (8, 13.5))
    elif name == "keyboard":
        _rr(cr, 1, 3.5, 14, 9, 1.5)
        cr.stroke()
        for x in (4, 6.7, 9.3, 12):
            cr.rectangle(x - 0.5, 6, 1, 1)
        cr.fill()
        L((4.5, 10), (11.5, 10))


class Icon(Gtk.DrawingArea):
    """Monochrome glyph icon."""

    def __init__(self, name, size=16, color=None):
        super().__init__()
        self.name = name
        self.size = size
        self.color = color
        self.set_size_request(size, size)
        self.set_valign(Gtk.Align.CENTER)
        self.connect("draw", self._draw)

    def set_icon(self, name):
        self.name = name
        self.queue_draw()

    def _draw(self, w, cr):
        col = self.color or (1, 1, 1, 1)
        p = self.get_parent()
        if not self.is_sensitive() or (p and not p.is_sensitive()):
            col = (1, 1, 1, 0.36)
        draw_glyph(cr, self.name, self.size, col)


class WinTitlebar(Gtk.HeaderBar):
    """Title bar with Windows 11 caption buttons. Pack your own widgets with pack_start()."""

    def __init__(self, window):
        super().__init__()
        self.win = window
        self.get_style_context().add_class("fx-title")
        self.set_show_close_button(False)
        self.set_custom_title(Gtk.Box())
        wc = Gtk.Box(spacing=0)
        wc.set_valign(Gtk.Align.START)
        for name, cb, extra in (("min", lambda *_: window.iconify(), ""),
                                ("max", lambda *_: self.toggle_max(), ""),
                                ("close", lambda *_: window.close(), "close")):
            b = Gtk.Button()
            b.get_style_context().add_class("wc")
            if extra:
                b.get_style_context().add_class(extra)
            ic = Icon(name, 10)
            b.add(ic)
            b.connect("clicked", cb)
            if name == "max":
                self.max_icon = ic
            wc.pack_start(b, False, False, 0)
        self.pack_end(wc)
        window.connect("window-state-event", self._wstate)
        window.set_titlebar(self)

    def toggle_max(self):
        if self.win.is_maximized():
            self.win.unmaximize()
        else:
            self.win.maximize()

    def _wstate(self, w, ev):
        maxed = bool(ev.new_window_state & Gdk.WindowState.MAXIMIZED)
        self.max_icon.set_icon("restore" if maxed else "max")
        return False


def app_title(text, logo_draw):
    """Little logo + app name for the left of the title bar."""
    box = Gtk.Box(spacing=8)
    box.set_margin_start(14)
    da = Gtk.DrawingArea()
    da.set_size_request(16, 16)
    da.set_valign(Gtk.Align.CENTER)
    da.connect("draw", lambda w, cr: logo_draw(cr, 16))
    box.pack_start(da, False, False, 0)
    lb = Gtk.Label(label=text)
    lb.get_style_context().add_class("app-title")
    box.pack_start(lb, False, False, 0)
    return box, lb


class JsonSettings(dict):
    def __init__(self, app_dir, defaults):
        super().__init__(defaults)
        self.dir = os.path.join(HOME, ".config", app_dir)
        self.file = os.path.join(self.dir, "settings.json")
        try:
            with open(self.file) as f:
                self.update(json.load(f))
        except (OSError, ValueError):
            pass

    def save(self):
        try:
            os.makedirs(self.dir, exist_ok=True)
            tmp = self.file + ".tmp"
            with open(tmp, "w") as f:
                json.dump(self, f, indent=1)
            os.replace(tmp, self.file)
        except OSError:
            pass


def vsep():
    s = Gtk.Box()
    s.get_style_context().add_class("vsep")
    return s


def logo_pixbuf(draw, size):
    surf = cairo.ImageSurface(cairo.FORMAT_ARGB32, size, size)
    draw(cairo.Context(surf), size)
    return Gdk.pixbuf_get_from_surface(surf, 0, 0, size, size)
def label(text="", cls=None, xalign=0.0, ellipsize=False, wrap=False, markup=False):
    lb = Gtk.Label()
    if markup:
        lb.set_markup(text)
    else:
        lb.set_text(text)
    lb.set_xalign(xalign)
    if cls:
        for c in cls.split():
            lb.get_style_context().add_class(c)
    if ellipsize:
        lb.set_ellipsize(Pango.EllipsizeMode.END)
    if wrap:
        lb.set_line_wrap(True)
        lb.set_line_wrap_mode(Pango.WrapMode.WORD_CHAR)
    return lb


def add_cls(w, *cls):
    ctx = w.get_style_context()
    for c in cls:
        ctx.add_class(c)
    return w


def tb_button(icon, text, cb=None, chev=False, tip=None):
    b = Gtk.Button()
    add_cls(b, "tb")
    box = Gtk.Box(spacing=7)
    if icon:
        box.pack_start(Icon(icon), False, False, 0)
    if text:
        box.pack_start(Gtk.Label(label=text), False, False, 0)
    if chev:
        box.pack_start(Icon("chevdown", 10), False, False, 0)
    b.add(box)
    b.set_relief(Gtk.ReliefStyle.NONE)
    b.set_tooltip_text(tip or text or None)
    if cb:
        b.connect("clicked", lambda w: cb(w))
    return b


def icon_button(icon, cb, cls="nb", tip=None, size=16):
    b = Gtk.Button()
    add_cls(b, cls)
    b.add(Icon(icon, size))
    b.set_relief(Gtk.ReliefStyle.NONE)
    if tip:
        b.set_tooltip_text(tip)
    b.connect("clicked", lambda w: cb(w))
    return b


def build_menu(spec):
    """spec: list of None (separator) or dicts: label, cb, icon, place, accel, sensitive, checked, sub."""
    menu = Gtk.Menu()
    add_cls(menu, "context-menu")
    for it in spec:
        if it is None:
            menu.append(Gtk.SeparatorMenuItem())
            continue
        mi = Gtk.MenuItem()
        box = Gtk.Box(spacing=12)
        ic = None
        if it.get("checked"):
            ic = Icon("check", 16)
        elif it.get("pixbuf") is not None:
            ic = Gtk.Image.new_from_pixbuf(it["pixbuf"])
        elif it.get("icon"):
            ic = Icon(it["icon"], 16)
        if ic is None:
            ic = Gtk.Box()
            ic.set_size_request(16, 16)
        box.pack_start(ic, False, False, 0)
        lb = Gtk.Label(label=it["label"])
        lb.set_xalign(0)
        box.pack_start(lb, True, True, 0)
        if it.get("accel"):
            al = Gtk.Label(label=it["accel"])
            al.get_style_context().add_class("dim-label")
            al.set_opacity(0.55)
            al.set_margin_start(24)
            box.pack_end(al, False, False, 0)
        mi.add(box)
        mi.set_sensitive(it.get("sensitive", True))
        if it.get("sub") is not None:
            mi.set_submenu(build_menu(it["sub"]))
        elif it.get("cb"):
            cb = it["cb"]
            mi.connect("activate", lambda w, cb=cb: GLib.idle_add(cb))
        menu.append(mi)
    menu.show_all()
    return menu


def popup_menu(menu, widget, event=None):
    menu.attach_to_widget(widget, None)
    if event is not None:
        menu.popup_at_pointer(event)
    else:
        menu.popup_at_widget(widget, Gdk.Gravity.SOUTH_WEST, Gdk.Gravity.NORTH_WEST, None)


# ----------------------------------------------------------------------------
# Windows 11 style dialogs
# ----------------------------------------------------------------------------
class WinDialog(Gtk.Window):
    def __init__(self, parent, title, body=None, primary="OK", secondary="Cancel", width=440, extra=None,
                 markup=False):
        super().__init__(transient_for=parent, modal=True)
        add_cls(self, "fx-dialog")
        self.set_decorated(False)
        self.set_resizable(False)
        self.set_default_size(width, -1)
        self.set_position(Gtk.WindowPosition.CENTER_ON_PARENT)
        self.set_type_hint(Gdk.WindowTypeHint.DIALOG)
        self.set_title(title)
        self.cb = None
        outer = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.body = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        add_cls(self.body, "dlg-body")
        self.body.pack_start(label(title, "dlg-title"), False, False, 0)
        if body:
            lb = label(body, wrap=True, markup=markup)
            lb.set_max_width_chars(60)
            self.body.pack_start(lb, False, False, 0)
        if extra:
            self.body.pack_start(extra, False, False, 0)
        outer.pack_start(self.body, True, True, 0)
        self.foot = Gtk.Box(spacing=8, homogeneous=True)
        add_cls(self.foot, "dlg-foot")
        self.primary_btn = None
        if primary:
            pb = Gtk.Button(label=primary)
            add_cls(pb, "primary")
            pb.connect("clicked", lambda *_: self.done(True))
            self.foot.pack_start(pb, True, True, 0)
            pb.set_can_default(True)
            self.set_default(pb)
            self.primary_btn = pb
        if secondary:
            sb = Gtk.Button(label=secondary)
            add_cls(sb, "secondary")
            sb.connect("clicked", lambda *_: self.done(False))
            self.foot.pack_start(sb, True, True, 0)
        if primary or secondary:
            outer.pack_start(self.foot, False, False, 0)
        self.add(outer)
        self.connect("key-press-event", self._key)

    def _key(self, w, ev):
        if ev.keyval == Gdk.KEY_Escape:
            self.done(False)
            return True
        return False

    def done(self, ok):
        cb = self.cb
        self.cb = None
        self.hide()
        if cb:
            cb(ok)
        self.destroy()

    def ask(self, cb=None):
        self.cb = cb
        self.show_all()
        if self.primary_btn:
            self.primary_btn.grab_focus()


def info_dialog(parent, title, body):
    WinDialog(parent, title, body, primary="OK", secondary=None).ask()


def confirm(parent, title, body, yes, cb, no="Cancel"):
    WinDialog(parent, title, body, primary=yes, secondary=no).ask(cb)


def error_list_dialog(parent, title, errors):
    lines = []
    for p, msg in errors[:12]:
        lines.append(f"{os.path.basename(p) or p}: {msg}" if p else msg)
    if len(errors) > 12:
        lines.append(f"...and {len(errors) - 12} more")
    info_dialog(parent, title, "\n".join(lines))




# ----------------------------------------------------------------------------
# PowerPoint: slide model, drawing, files
# ----------------------------------------------------------------------------
import io  # noqa: E402
import copy  # noqa: E402
import datetime  # noqa: E402
gi.require_version("PangoCairo", "1.0")
from gi.repository import PangoCairo  # noqa: E402

EMU = 12700  # per point
FONT_SUB = {"Calibri": "Carlito", "Calibri Light": "Carlito", "Cambria": "Caladea", "Arial": "Liberation Sans",
            "Times New Roman": "Liberation Serif", "Courier New": "Liberation Mono", "Aptos": "Carlito",
            "Aptos Display": "Carlito", "Segoe UI": "Noto Sans", "Century Gothic": "DejaVu Sans"}
FONTS = ["Calibri", "Calibri Light", "Cambria", "Arial", "Times New Roman", "Georgia", "Verdana", "Courier New",
         "DejaVu Sans", "Liberation Sans", "Noto Sans", "Open Sans"]
SIZES = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 88]

THEMES = {
    "office": {"name": "Office", "bg": "#FFFFFF", "title": "#000000", "text": "#262626", "accent": "#4472C4",
               "accent2": "#ED7D31", "tfont": "Calibri Light", "bfont": "Calibri"},
    "dark": {"name": "Dark", "bg": "#1E1E1E", "title": "#FFFFFF", "text": "#E6E6E6", "accent": "#3FA9F5",
             "accent2": "#F5A623", "tfont": "Calibri Light", "bfont": "Calibri"},
    "slate": {"name": "Slate", "bg": "#2E3A46", "title": "#FFFFFF", "text": "#DCE3EA", "accent": "#F5A623",
              "accent2": "#5BC0BE", "tfont": "Calibri", "bfont": "Calibri"},
    "ion": {"name": "Ion", "bg": "#0F5E63", "title": "#FFFFFF", "text": "#E8F4F4", "accent": "#F2C14E",
            "accent2": "#F78154", "tfont": "Calibri Light", "bfont": "Calibri"},
    "paper": {"name": "Paper", "bg": "#F4F1EA", "title": "#3A3A3A", "text": "#4A4A4A", "accent": "#C0504D",
              "accent2": "#4F81BD", "tfont": "Cambria", "bfont": "Calibri"},
    "berry": {"name": "Berry", "bg": "#3B1F3F", "title": "#FFFFFF", "text": "#F3E6F4", "accent": "#E85D75",
              "accent2": "#F7B267", "tfont": "Calibri Light", "bfont": "Calibri"},
    "fresh": {"name": "Fresh", "bg": "#FFFFFF", "title": "#1E5631", "text": "#333333", "accent": "#4C9A2A",
              "accent2": "#A4DE02", "tfont": "Calibri", "bfont": "Calibri"},
    "sky": {"name": "Sky", "bg": "#EAF4FB", "title": "#0B3954", "text": "#1D3D52", "accent": "#087E8B",
            "accent2": "#FF5A5F", "tfont": "Calibri Light", "bfont": "Calibri"},
}
LAYOUTS = [("title", "Title Slide"), ("title_content", "Title and Content"), ("section", "Section Header"),
           ("two", "Two Content"), ("title_only", "Title Only"), ("blank", "Blank")]
PROMPTS = {"title": "Click to add title", "subtitle": "Click to add subtitle", "body": "Click to add text"}
SHAPES = [("rect", "Rectangle"), ("round", "Rounded rectangle"), ("ellipse", "Oval"), ("triangle", "Triangle"),
          ("arrow", "Right arrow"), ("star", "Star"), ("line", "Line")]
TEXT_COLORS = ["#000000", "#FFFFFF", "#7F7F7F", "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#00B050", "#0070C0",
               "#002060", "#7030A0"]


def font_for(name):
    return FONT_SUB.get(name, name)


def text_shape(role, x, y, w, h, size, align="left", valign="top", bullets=False):
    return {"kind": "text", "role": role, "x": x, "y": y, "w": w, "h": h, "text": "", "font": None, "size": size,
            "b": False, "i": False, "u": False, "color": None, "align": align, "valign": valign, "bullets": bullets,
            "fill": None, "line": None, "autofit": role == "body"}


def layout_shapes(layout, W, H):
    k = W / 960.0
    X = lambda v: v * k  # noqa: E731
    if layout == "title":
        return [text_shape("title", X(72), 150, X(816), 130, 54, "center", "bottom"),
                text_shape("subtitle", X(72), 292, X(816), 90, 24, "center", "top")]
    if layout == "title_content":
        return [text_shape("title", X(48), 28, X(864), 86, 40, "left", "middle"),
                text_shape("body", X(48), 128, X(864), 370, 24, "left", "top", True)]
    if layout == "section":
        return [text_shape("title", X(64), 190, X(832), 120, 48, "left", "bottom"),
                text_shape("subtitle", X(64), 318, X(832), 64, 20, "left", "top")]
    if layout == "two":
        return [text_shape("title", X(48), 28, X(864), 86, 40, "left", "middle"),
                text_shape("body", X(48), 128, X(420), 370, 22, "left", "top", True),
                text_shape("body", X(492), 128, X(420), 370, 22, "left", "top", True)]
    if layout == "title_only":
        return [text_shape("title", X(48), 28, X(864), 86, 40, "left", "middle")]
    return []


def new_slide(layout, W, H):
    return {"layout": layout, "bg": None, "shapes": layout_shapes(layout, W, H), "notes": ""}


def new_pres():
    p = {"w": 960.0, "h": 540.0, "theme": "office", "fade": True, "slides": []}
    p["slides"].append(new_slide("title", p["w"], p["h"]))
    return p


def hex_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def darker(h, f=0.75):
    r, g, b = hex_rgb(h)
    return "#%02X%02X%02X" % (int(r * 255 * f), int(g * 255 * f), int(b * 255 * f))


def luminance(h):
    r, g, b = hex_rgb(h)
    return 0.299 * r + 0.587 * g + 0.114 * b


# ---------------------------------------------------------------- drawing
_PB_CACHE = {}


def pixbuf_for(data):
    key = id(data)
    hit = _PB_CACHE.get(key)
    if hit is not None and hit[0] is data:
        return hit[1]
    try:
        loader = GdkPixbuf.PixbufLoader()
        loader.write(data)
        loader.close()
        pb = loader.get_pixbuf()
    except GLib.Error:
        pb = None
    if len(_PB_CACHE) > 200:
        _PB_CACHE.clear()
    _PB_CACHE[key] = (data, pb)
    return pb


def shape_path(cr, s):
    x, y, w, h = s["x"], s["y"], s["w"], s["h"]
    k = s["kind"]
    cr.new_path()
    if k == "ellipse":
        cr.save()
        cr.translate(x + w / 2, y + h / 2)
        cr.scale(max(0.01, abs(w) / 2), max(0.01, abs(h) / 2))
        cr.arc(0, 0, 1, 0, 2 * math.pi)
        cr.restore()
    elif k == "round":
        _rr(cr, x, y, w, h, min(abs(w), abs(h)) * 0.16)
    elif k == "triangle":
        cr.move_to(x + w / 2, y)
        cr.line_to(x + w, y + h)
        cr.line_to(x, y + h)
        cr.close_path()
    elif k == "arrow":
        hw_ = min(w * 0.4, h)
        cr.move_to(x, y + h * 0.25)
        cr.line_to(x + w - hw_, y + h * 0.25)
        cr.line_to(x + w - hw_, y)
        cr.line_to(x + w, y + h / 2)
        cr.line_to(x + w - hw_, y + h)
        cr.line_to(x + w - hw_, y + h * 0.75)
        cr.line_to(x, y + h * 0.75)
        cr.close_path()
    elif k == "star":
        cx, cy = x + w / 2, y + h / 2
        for i in range(10):
            a = -math.pi / 2 + i * math.pi / 5
            rr = 1.0 if i % 2 == 0 else 0.4
            px, py = cx + math.cos(a) * w / 2 * rr, cy + math.sin(a) * h / 2 * rr
            (cr.move_to if i == 0 else cr.line_to)(px, py)
        cr.close_path()
    else:
        cr.rectangle(x, y, w, h)


def para_layouts(cr, s, theme, scale_font=1.0):
    """Pango layouts for each paragraph of a text-bearing shape: list of (layout, indent, bullet_layout)."""
    role = s.get("role")
    fam = font_for(s.get("font") or (theme["tfont"] if role == "title" else theme["bfont"]))
    default_col = theme["title"] if role == "title" else theme["text"]
    if s["kind"] not in ("text",):
        fill = s.get("fill", theme["accent"])
        default_col = "#FFFFFF" if fill and luminance(fill) < 0.6 else "#000000"
    col = s.get("color") or default_col
    inset = 7.2
    out = []
    for raw in s["text"].split("\n"):
        level = len(raw) - len(raw.lstrip("\t")) if s.get("bullets") else 0
        text = raw.lstrip("\t") if s.get("bullets") else raw
        size = s["size"] * scale_font * (1 - 0.1 * min(level, 3))
        lay = PangoCairo.create_layout(cr)
        d = Pango.FontDescription()
        d.set_family(fam)
        d.set_absolute_size(size * Pango.SCALE)
        if s.get("b"):
            d.set_weight(Pango.Weight.BOLD)
        if s.get("i"):
            d.set_style(Pango.Style.ITALIC)
        lay.set_font_description(d)
        attrs = Pango.AttrList()
        if s.get("u"):
            attrs.insert(Pango.attr_underline_new(Pango.Underline.SINGLE))
        lay.set_attributes(attrs)
        indent = 0.0
        bl = None
        if s.get("bullets") and text != "":
            indent = size * 0.9 + level * 36
            bl = PangoCairo.create_layout(cr)
            bl.set_font_description(d)
            bl.set_text("•" if level % 2 == 0 else "–", -1)
        lay.set_width(int(max(1, abs(s["w"]) - 2 * inset - indent) * Pango.SCALE))
        lay.set_wrap(Pango.WrapMode.WORD_CHAR)
        lay.set_alignment({"left": Pango.Alignment.LEFT, "center": Pango.Alignment.CENTER,
                           "right": Pango.Alignment.RIGHT}.get(s.get("align", "left"), Pango.Alignment.LEFT))
        lay.set_text(text if text else " ", -1)
        lay.set_spacing(int(size * 0.1 * Pango.SCALE))
        out.append((lay, indent, bl, col, size))
    return out


def text_height(pls):
    tot = 0.0
    for i, (lay, ind, bl, col, size) in enumerate(pls):
        tot += lay.get_extents()[1].height / Pango.SCALE
        if i:
            tot += size * 0.25
    return tot


def draw_text(cr, s, theme, prompt=None):
    inset_x, inset_y = 7.2, 3.6
    x, y, h = s["x"], s["y"], abs(s["h"])
    if not s["text"] and prompt:
        tmp = dict(s, text=prompt, color="#8C8C8C" if luminance(theme["bg"]) > 0.5 else "#A0A0A0", bullets=False)
        pls = para_layouts(cr, tmp, theme)
    elif not s["text"]:
        return
    else:
        scale = 1.0
        pls = para_layouts(cr, s, theme)
        if s.get("autofit"):
            while text_height(pls) > h - 2 * inset_y and scale > 0.45:
                scale -= 0.05
                pls = para_layouts(cr, s, theme, scale)
    th = text_height(pls)
    va = s.get("valign", "top")
    ty = y + inset_y if va == "top" else (y + h - inset_y - th if va == "bottom" else y + (h - th) / 2)
    cr.save()
    for i, (lay, ind, bl, col, size) in enumerate(pls):
        if i:
            ty += size * 0.25
        cr.set_source_rgb(*hex_rgb(col))
        if bl is not None:
            lvl_off = ind - size * 0.9
            cr.move_to(x + inset_x + lvl_off, ty)
            PangoCairo.show_layout(cr, bl)
        cr.move_to(x + inset_x + ind, ty)
        PangoCairo.show_layout(cr, lay)
        ty += lay.get_extents()[1].height / Pango.SCALE
    cr.restore()


def draw_table(cr, s, theme):
    rows = s.get("rows") or [[""]]
    nr, nc = len(rows), max(len(r) for r in rows)
    x, y, w, h = s["x"], s["y"], s["w"], s["h"]
    cw, rh = w / nc, h / nr
    acc = theme["accent"]
    for i in range(nr):
        for j in range(nc):
            if i == 0:
                cr.set_source_rgb(*hex_rgb(acc))
            else:
                cr.set_source_rgb(*hex_rgb("#E9EDF4" if i % 2 else "#D0D8E8")) if luminance(theme["bg"]) > 0.5 else \
                    cr.set_source_rgba(1, 1, 1, 0.12 if i % 2 else 0.06)
            cr.rectangle(x + j * cw, y + i * rh, cw, rh)
            cr.fill()
    cr.set_source_rgb(1, 1, 1) if luminance(theme["bg"]) > 0.5 else cr.set_source_rgba(1, 1, 1, 0.3)
    cr.set_line_width(1)
    for i in range(1, nr):
        cr.move_to(x, y + i * rh)
        cr.line_to(x + w, y + i * rh)
    for j in range(1, nc):
        cr.move_to(x + j * cw, y)
        cr.line_to(x + j * cw, y + h)
    cr.stroke()
    lay = PangoCairo.create_layout(cr)
    for i, r in enumerate(rows):
        for j in range(nc):
            t = r[j] if j < len(r) else ""
            if not t:
                continue
            d = Pango.FontDescription()
            d.set_family(font_for(theme["bfont"]))
            d.set_absolute_size(s.get("size", 16) * Pango.SCALE)
            if i == 0:
                d.set_weight(Pango.Weight.BOLD)
            lay.set_font_description(d)
            lay.set_width(int(max(1, cw - 10) * Pango.SCALE))
            lay.set_wrap(Pango.WrapMode.WORD_CHAR)
            lay.set_text(t, -1)
            tw, th = lay.get_pixel_size()
            if i == 0:
                cr.set_source_rgb(1, 1, 1)
            else:
                cr.set_source_rgb(*hex_rgb("#1A1A1A" if luminance(theme["bg"]) > 0.5 else "#F2F2F2"))
            cr.move_to(x + j * cw + 5, y + i * rh + (rh - th) / 2)
            PangoCairo.show_layout(cr, lay)


def render_slide(cr, pres, slide, edit=False, hide=None):
    """Draw a slide in slide points (cr already scaled)."""
    theme = THEMES.get(pres.get("theme"), THEMES["office"])
    W, H = pres["w"], pres["h"]
    cr.save()
    cr.rectangle(0, 0, W, H)
    cr.clip()
    cr.set_source_rgb(*hex_rgb(slide.get("bg") or theme["bg"]))
    cr.paint()
    for s in slide["shapes"]:
        if s is hide:
            if s["kind"] != "text":
                _draw_shape_body(cr, s, theme)
            continue
        k = s["kind"]
        if k == "text":
            if s.get("fill"):
                cr.set_source_rgb(*hex_rgb(s["fill"]))
                cr.rectangle(s["x"], s["y"], s["w"], s["h"])
                cr.fill()
            if s.get("line"):
                cr.set_source_rgb(*hex_rgb(s["line"]))
                cr.set_line_width(s.get("line_w", 1.5))
                cr.rectangle(s["x"], s["y"], s["w"], s["h"])
                cr.stroke()
            prompt = PROMPTS.get(s.get("role")) if edit and s.get("role") else None
            if edit and s.get("role") and not s["text"]:
                cr.save()
                cr.set_source_rgba(0.55, 0.55, 0.55, 0.9)
                cr.set_line_width(0.8)
                cr.set_dash([4, 3])
                cr.rectangle(s["x"], s["y"], s["w"], s["h"])
                cr.stroke()
                cr.restore()
            draw_text(cr, s, theme, prompt)
        elif k == "image":
            pb = pixbuf_for(s["img"])
            if pb is not None:
                cr.save()
                cr.rectangle(s["x"], s["y"], s["w"], s["h"])
                cr.clip()
                cr.translate(s["x"], s["y"])
                cr.scale(s["w"] / pb.get_width(), s["h"] / pb.get_height())
                Gdk.cairo_set_source_pixbuf(cr, pb, 0, 0)
                cr.get_source().set_filter(cairo.FILTER_GOOD)
                cr.paint()
                cr.restore()
        elif k == "table":
            draw_table(cr, s, theme)
        else:
            _draw_shape_body(cr, s, theme)
            if s.get("text"):
                draw_text(cr, dict(s, valign=s.get("valign", "middle"), align=s.get("align", "center")), theme)
    cr.restore()


def _draw_shape_body(cr, s, theme):
    cr.new_path()
    if s["kind"] == "line":
        cr.set_source_rgb(*hex_rgb(s.get("line") or theme["accent"]))
        cr.set_line_width(s.get("line_w", 2.25))
        cr.set_line_cap(cairo.LINE_CAP_ROUND)
        cr.move_to(s["x"], s["y"])
        cr.line_to(s["x"] + s["w"], s["y"] + s["h"])
        cr.stroke()
        return
    shape_path(cr, s)
    fill = s.get("fill", theme["accent"])
    if fill:
        cr.set_source_rgb(*hex_rgb(fill))
        cr.fill_preserve()
    line = s.get("line", darker(theme["accent"]))
    if line:
        cr.set_source_rgb(*hex_rgb(line))
        cr.set_line_width(s.get("line_w", 1.0))
        cr.stroke()
    else:
        cr.new_path()


# ---------------------------------------------------------------- files
def _rgb_of(color_format):
    try:
        if color_format is not None and color_format.type is not None and color_format.rgb is not None:
            return "#" + str(color_format.rgb)
    except (AttributeError, ValueError, TypeError):
        pass
    return None


def load_pptx(path):
    from pptx import Presentation
    from pptx.util import Emu  # noqa: F401
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    prs = Presentation(path)
    W, H = prs.slide_width / EMU, prs.slide_height / EMU
    pres = {"w": W, "h": H, "theme": "office", "fade": True, "slides": []}
    cat = prs.core_properties.category or ""
    if cat.startswith("pi-powerpoint:"):
        for kv in cat.split(":", 1)[1].split(";"):
            k, _, v = kv.partition("=")
            if k == "theme" and v in THEMES:
                pres["theme"] = v
            if k == "fade":
                pres["fade"] = v == "1"
    theme = THEMES[pres["theme"]]

    def conv_text(sh, s):
        tf = sh.text_frame
        paras = []
        first_run = None
        al = None
        for p in tf.paragraphs:
            t = "".join(r.text for r in p.runs) if p.runs else p.text
            lvl = p.level or 0
            paras.append("\t" * lvl + t)
            if first_run is None and p.runs:
                first_run = p.runs[0]
            if al is None and p.alignment is not None:
                al = p.alignment
        s["text"] = "\n".join(paras).replace("\x0b", " ")
        if not s["text"].strip():
            s["text"] = ""
        if first_run is not None:
            f = first_run.font
            if f.size:
                s["size"] = f.size.pt
            if f.bold:
                s["b"] = True
            if f.italic:
                s["i"] = True
            if f.underline:
                s["u"] = True
            if f.name:
                s["font"] = f.name
            c = _rgb_of(f.color) if f.color and f.color.type is not None else None
            if c:
                s["color"] = c
        if al is not None:
            s["align"] = {PP_ALIGN.CENTER: "center", PP_ALIGN.RIGHT: "right"}.get(al, "left")
        try:
            anc = tf.vertical_anchor
            if anc is not None:
                s["valign"] = {MSO_ANCHOR.MIDDLE: "middle", MSO_ANCHOR.BOTTOM: "bottom"}.get(anc, "top")
        except (AttributeError, ValueError):
            pass
        # bullets written as <a:buChar> or inherited in body placeholders
        xml = sh._element.xml
        if "<a:buChar" in xml:
            s["bullets"] = True
        if "<a:buNone" in xml and "<a:buChar" not in xml:
            s["bullets"] = False

    def conv(sh, dx=0.0, dy=0.0):
        out = []
        try:
            x, y = (sh.left or 0) / EMU + dx, (sh.top or 0) / EMU + dy
            w, h = (sh.width or 0) / EMU, (sh.height or 0) / EMU
        except (AttributeError, ValueError):
            return out
        st = sh.shape_type
        if st == MSO_SHAPE_TYPE.GROUP:
            for ch in sh.shapes:
                out += conv(ch, 0, 0)
            return out
        if st == MSO_SHAPE_TYPE.PICTURE or (sh.is_placeholder and hasattr(sh, "image")):
            try:
                return [{"kind": "image", "x": x, "y": y, "w": w, "h": h, "img": sh.image.blob}]
            except (AttributeError, ValueError):
                return []
        if getattr(sh, "has_table", False) and sh.has_table:
            rows = [[c.text for c in r.cells] for r in sh.table.rows]
            return [{"kind": "table", "x": x, "y": y, "w": w, "h": h, "rows": rows, "size": 16}]
        if st == MSO_SHAPE_TYPE.LINE or sh.__class__.__name__ == "Connector":
            try:
                x1, y1 = sh.begin_x / EMU, sh.begin_y / EMU
                x2, y2 = sh.end_x / EMU, sh.end_y / EMU
            except AttributeError:
                x1, y1, x2, y2 = x, y, x + w, y + h
            s = {"kind": "line", "x": x1, "y": y1, "w": x2 - x1, "h": y2 - y1, "line": None, "line_w": 2.25}
            try:
                s["line"] = _rgb_of(sh.line.color)
                if sh.line.width:
                    s["line_w"] = sh.line.width.pt
            except (AttributeError, ValueError, TypeError):
                pass
            return [s]
        role = None
        if sh.is_placeholder:
            pt = sh.placeholder_format.type
            if pt in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.VERTICAL_TITLE):
                role = "title"
            elif pt == PP_PLACEHOLDER.SUBTITLE:
                role = "subtitle"
            elif pt in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.VERTICAL_BODY):
                role = "body"
            elif pt in (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER):
                if not sh.has_text_frame or not sh.text_frame.text.strip():
                    return []
        if role is None and st == MSO_SHAPE_TYPE.TEXT_BOX:
            nm = (sh.name or "").split(" ")[0]
            role = {"Title": "title", "Subtitle": "subtitle", "Content": "body"}.get(nm)
        kind = "text"
        if st == MSO_SHAPE_TYPE.AUTO_SHAPE and not sh.is_placeholder:
            try:
                ast = str(sh.auto_shape_type).split(".")[-1].split(" ")[0].upper()
            except (NotImplementedError, ValueError, AttributeError):
                ast = "RECTANGLE"
            kind = {"OVAL": "ellipse", "ROUNDED_RECTANGLE": "round", "ISOSCELES_TRIANGLE": "triangle",
                    "RIGHT_ARROW": "arrow", "STAR_5_POINT": "star"}.get(ast, "rect")
            if ast == "RECTANGLE" or ast == "TEXT_BOX":
                kind = "rect"
        defaults = {"title": 40, "subtitle": 22, "body": 24}
        s = text_shape(role, x, y, w, h, defaults.get(role, 18), "left", "top", role == "body")
        s["kind"] = kind
        if role == "title":
            s["valign"] = "middle"
        if sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.CENTER_TITLE:
            s["align"], s["valign"], s["size"] = "center", "bottom", 44
        elif role == "subtitle":
            s["align"] = "center"
        if kind != "text":
            s["align"], s["valign"] = "center", "middle"
            s["fill"] = theme["accent"]
            s["line"] = darker(theme["accent"])
        try:
            fill = sh.fill
            if fill.type == 1:  # solid
                s["fill"] = _rgb_of(fill.fore_color) or s["fill"]
            elif fill.type == 5 and kind != "text":  # background / none
                s["fill"] = None
        except (AttributeError, TypeError, NotImplementedError):
            pass
        try:
            if sh.line.fill.type == 1:
                s["line"] = _rgb_of(sh.line.color) or s.get("line")
            elif kind != "text" and sh.line.fill.type == 5:
                s["line"] = None
        except (AttributeError, TypeError, NotImplementedError):
            pass
        if sh.has_text_frame:
            conv_text(sh, s)
            if kind == "text" and not role and not s["text"] and not s.get("fill"):
                return []
        return [s]

    for sl in prs.slides:
        slide = {"layout": "blank", "bg": None, "shapes": [], "notes": ""}
        try:
            bgf = sl.background.fill
            if bgf.type == 1:
                slide["bg"] = _rgb_of(bgf.fore_color)
        except (AttributeError, TypeError, NotImplementedError):
            pass
        for sh in sl.shapes:
            try:
                slide["shapes"] += conv(sh)
            except Exception:
                pass
        if sl.has_notes_slide:
            try:
                slide["notes"] = sl.notes_slide.notes_text_frame.text
            except AttributeError:
                pass
        pres["slides"].append(slide)
    if not pres["slides"]:
        pres["slides"].append(new_slide("title", W, H))
    return pres


def save_pptx(pres, path):
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.oxml.ns import qn
    from lxml import etree
    theme = THEMES.get(pres["theme"], THEMES["office"])
    prs = Presentation()
    prs.slide_width = Emu(int(pres["w"] * EMU))
    prs.slide_height = Emu(int(pres["h"] * EMU))
    prs.core_properties.category = f"pi-powerpoint:theme={pres['theme']};fade={1 if pres.get('fade') else 0}"
    blank = prs.slide_layouts[6]
    E = lambda v: Emu(int(round(v * EMU)))  # noqa: E731
    rgb = lambda h: RGBColor.from_string(h.lstrip("#").upper())  # noqa: E731

    def set_bullet(p, on, level, size):
        pPr = p._p.get_or_add_pPr()
        for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buFont"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)
        if on:
            ind = int((size * 0.9 + level * 36) * EMU)
            pPr.set("marL", str(ind))
            pPr.set("indent", str(-int(size * 0.9 * EMU)))
            bf = etree.SubElement(pPr, qn("a:buFont"))
            bf.set("typeface", "Arial")
            bc = etree.SubElement(pPr, qn("a:buChar"))
            bc.set("char", "•" if level % 2 == 0 else "–")
        else:
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            etree.SubElement(pPr, qn("a:buNone"))

    def fill_text(tf, s, default_color):
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = {"middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}.get(s.get("valign"), MSO_ANCHOR.TOP)
        if s.get("autofit"):
            bp = tf._txBody.find(qn("a:bodyPr"))
            for el in list(bp):
                if el.tag in (qn("a:spAutoFit"), qn("a:noAutofit"), qn("a:normAutofit")):
                    bp.remove(el)
            etree.SubElement(bp, qn("a:normAutofit"))
        lines = s["text"].split("\n")
        for i, raw in enumerate(lines):
            level = len(raw) - len(raw.lstrip("\t")) if s.get("bullets") else 0
            text = raw.lstrip("\t") if s.get("bullets") else raw
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(s.get("align"), PP_ALIGN.LEFT)
            if level:
                p.level = min(level, 8)
            size = s["size"] * (1 - 0.1 * min(level, 3))
            set_bullet(p, bool(s.get("bullets")) and text != "", level, size)
            r = p.add_run()
            r.text = text
            f = r.font
            f.size = Pt(round(size * 2) / 2)
            f.bold = bool(s.get("b"))
            f.italic = bool(s.get("i"))
            f.underline = bool(s.get("u"))
            f.name = s.get("font") or (theme["tfont"] if s.get("role") == "title" else theme["bfont"])
            f.color.rgb = rgb(s.get("color") or default_color)

    for slide in pres["slides"]:
        sl = prs.slides.add_slide(blank)
        bg = sl.background.fill
        bg.solid()
        bg.fore_color.rgb = rgb(slide.get("bg") or theme["bg"])
        for s in slide["shapes"]:
            k = s["kind"]
            x, y, w, h = s["x"], s["y"], s["w"], s["h"]
            if k == "image":
                try:
                    sl.shapes.add_picture(io.BytesIO(s["img"]), E(x), E(y), E(w), E(h))
                except Exception:
                    pass
                continue
            if k == "line":
                ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x), E(y), E(x + w), E(y + h))
                ln.line.color.rgb = rgb(s.get("line") or theme["accent"])
                ln.line.width = Pt(s.get("line_w", 2.25))
                continue
            if k == "table":
                rows = s.get("rows") or [[""]]
                nr, nc = len(rows), max(len(r) for r in rows)
                gf = sl.shapes.add_table(nr, nc, E(x), E(y), E(w), E(h))
                for i, r in enumerate(rows):
                    for j in range(nc):
                        cell = gf.table.cell(i, j)
                        cell.text = r[j] if j < len(r) else ""
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                run.font.size = Pt(s.get("size", 16))
                continue
            if k == "text":
                sp = sl.shapes.add_textbox(E(x), E(y), E(w), E(h))
                if s.get("role"):
                    sp.name = {"title": "Title", "subtitle": "Subtitle", "body": "Content"}[s["role"]] + f" {len(sl.shapes)}"
                if s.get("fill"):
                    sp.fill.solid()
                    sp.fill.fore_color.rgb = rgb(s["fill"])
                if s.get("line"):
                    sp.line.color.rgb = rgb(s["line"])
                    sp.line.width = Pt(s.get("line_w", 1.5))
                if not s["text"]:
                    continue
                fill_text(sp.text_frame, s, theme["title"] if s.get("role") == "title" else theme["text"])
                continue
            mso = {"rect": MSO_SHAPE.RECTANGLE, "round": MSO_SHAPE.ROUNDED_RECTANGLE, "ellipse": MSO_SHAPE.OVAL,
                   "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE, "arrow": MSO_SHAPE.RIGHT_ARROW,
                   "star": MSO_SHAPE.STAR_5_POINT}.get(k, MSO_SHAPE.RECTANGLE)
            sp = sl.shapes.add_shape(mso, E(x), E(y), E(w), E(h))
            fill = s.get("fill", theme["accent"])
            if fill:
                sp.fill.solid()
                sp.fill.fore_color.rgb = rgb(fill)
            else:
                sp.fill.background()
            line = s.get("line", darker(theme["accent"]))
            if line:
                sp.line.color.rgb = rgb(line)
                sp.line.width = Pt(s.get("line_w", 1.0))
            else:
                sp.line.fill.background()
            sp.shadow.inherit = False
            if s.get("text"):
                fill_text(sp.text_frame, dict(s, valign=s.get("valign", "middle"), align=s.get("align", "center")),
                          "#FFFFFF" if fill and luminance(fill) < 0.6 else "#000000")
        if slide.get("notes"):
            sl.notes_slide.notes_text_frame.text = slide["notes"]
    prs.save(path)


def render_pdf(pres, path):
    W, H = pres["w"], pres["h"]
    surf = cairo.PDFSurface(path, W, H)
    cr = cairo.Context(surf)
    for slide in pres["slides"]:
        render_slide(cr, pres, slide)
        cr.show_page()
    surf.finish()


# ----------------------------------------------------------------------------
# PowerPoint: window
# ----------------------------------------------------------------------------
APP_ID = "io.github.pipowerpoint"
DEFAULTS = {"width": 1280, "height": 800, "max": False, "recent": [], "notes": True}

CSS = """
.ribbon-tabs { background-color: #202020; padding: 0 8px; }
button.rtab { background: transparent; border: none; box-shadow: none; border-radius: 0; padding: 6px 12px; min-height: 26px; }
button.rtab label { font-size: 12px; color: rgba(255,255,255,0.8); }
button.rtab:hover { background-color: rgba(255,255,255,0.05); }
button.rtab.sel { border-bottom: 2px solid #F0703A; }
button.rtab.sel label { color: #ffffff; font-weight: 600; }
.ribbon { background-color: #2C2C2C; padding: 4px 8px; border-bottom: 1px solid rgba(0,0,0,0.4); min-height: 40px; }
.rsep { background-color: rgba(255,255,255,0.1); min-width: 1px; margin: 6px 6px; }
button.rb { background: transparent; border: none; box-shadow: none; border-radius: 4px; min-width: 30px; min-height: 30px; padding: 0 6px; }
button.rb:hover { background-color: rgba(255,255,255,0.07); }
button.rb.on { background-color: rgba(240,112,58,0.28); }
button.rb label { font-size: 12px; color: #ffffff; }
button.themeb { background: transparent; border: 2px solid transparent; box-shadow: none; padding: 1px; border-radius: 4px; }
button.themeb.on { border-color: #F0703A; }
.thumbs { background-color: #1F1F1F; }
.thumbs row { background: transparent; padding: 6px 6px 6px 2px; }
.thumbs row:selected { background-color: rgba(240,112,58,0.18); }
label.tnum { color: rgba(255,255,255,0.7); font-size: 11px; margin: 0 6px 0 8px; }
.canvas-bg { background-color: #1B1B1B; }
.notes { background-color: #262626; border-top: 1px solid rgba(0,0,0,0.5); }
textview.notesv, textview.notesv text { background-color: #262626; color: #e6e6e6; }
.status { background-color: #2C2C2C; }
combobox.fontbox button, combobox.sizebox button { min-height: 26px; padding: 0 6px; }
"""


def draw_logo(cr, s):
    cr.save()
    cr.scale(s / 16.0, s / 16.0)
    cr.arc(10, 8, 6, 0, 2 * math.pi)
    cr.set_source_rgb(0.95, 0.47, 0.25)
    cr.fill()
    cr.move_to(10, 8)
    cr.arc(10, 8, 6, -math.pi / 2, 0)
    cr.close_path()
    cr.set_source_rgb(0.99, 0.65, 0.40)
    cr.fill()
    g = cairo.LinearGradient(0, 0, 16, 16)
    g.add_color_stop_rgb(0, 0.85, 0.30, 0.13)
    g.add_color_stop_rgb(1, 0.62, 0.18, 0.06)
    _rr(cr, 1, 3.5, 9, 9, 1.4)
    cr.set_source(g)
    cr.fill()
    cr.set_source_rgb(1, 1, 1)
    cr.set_line_width(1.2)
    cr.set_line_join(cairo.LINE_JOIN_ROUND)
    cr.move_to(4, 10.6)
    cr.line_to(4, 5.4)
    cr.line_to(6, 5.4)
    cr.curve_to(8.2, 5.4, 8.2, 8.4, 6, 8.4)
    cr.line_to(4, 8.4)
    cr.stroke()
    cr.restore()


_orig_extra = draw_glyph_extra


def draw_glyph_extra(cr, name, L, pi):  # noqa: F811
    if name in ("al-left", "al-center", "al-right"):
        for w, y in ((12, 3.5), (8, 6.5), (12, 9.5), (8, 12.5)):
            x0 = {"al-left": 2, "al-center": 8 - w / 2, "al-right": 14 - w}[name]
            L((x0, y), (x0 + w, y))
    elif name == "bullets":
        for y in (4, 8, 12):
            cr.arc(3, y, 1, 0, 2 * pi)
            cr.fill()
            L((6, y), (14, y))
    elif name == "fcolor":
        L((4, 11), (8, 2.5), (12, 11))
        L((5.6, 7.8), (10.4, 7.8))
        cr.rectangle(2.5, 12.8, 11, 2)
        cr.fill()
    elif name == "fill":
        L((3, 8), (8, 3), (12.5, 7.5), (7.5, 12.5), (3, 8))
        L((6, 1.5), (8, 3))
        cr.rectangle(2.5, 13.6, 11, 1.8)
        cr.fill()
    elif name == "outline":
        _rr(cr, 2.5, 2.5, 11, 9, 1)
        cr.stroke()
        cr.rectangle(2.5, 13.6, 11, 1.8)
        cr.fill()
    elif name == "newslide":
        _rr(cr, 1.5, 3, 11, 8, 1)
        cr.stroke()
        L((13, 9), (13, 15))
        L((10, 12), (16, 12))
    elif name == "layout":
        _rr(cr, 1.5, 2.5, 13, 11, 1)
        cr.stroke()
        L((1.5, 6), (14.5, 6))
        L((8, 6), (8, 13.5))
    elif name == "dup":
        _rr(cr, 4.5, 1.5, 10, 7.5, 1)
        cr.stroke()
        _rr(cr, 1.5, 6.5, 10, 7.5, 1)
        cr.stroke()
    elif name == "textbox":
        _rr(cr, 1.5, 2.5, 13, 11, 1)
        cr.stroke()
        L((5, 5.5), (11, 5.5))
        L((8, 5.5), (8, 11))
    elif name == "shapes":
        cr.rectangle(1.5, 7, 7.5, 7.5)
        cr.stroke()
        cr.arc(10.5, 5.5, 4, 0, 2 * pi)
        cr.stroke()
    elif name == "table":
        _rr(cr, 2, 3, 12, 10, 1)
        cr.stroke()
        L((2, 6.5), (14, 6.5))
        L((2, 10), (14, 10))
        L((6, 3), (6, 13))
        L((10, 3), (10, 13))
    elif name == "theme":
        _rr(cr, 1.5, 2.5, 13, 11, 1)
        cr.stroke()
        cr.rectangle(1.5, 10.5, 13, 3)
        cr.fill()
        cr.set_font_size(6)
        cr.move_to(4, 8.5)
        cr.show_text("Aa")
    elif name == "tofront":
        cr.rectangle(5.5, 5.5, 8, 8)
        cr.fill()
        cr.rectangle(2.5, 2.5, 8, 8)
        cr.stroke()
    elif name == "toback":
        cr.rectangle(2.5, 2.5, 8, 8)
        cr.fill()
        cr.set_source_rgb(0.17, 0.17, 0.17)
        cr.rectangle(5.5, 5.5, 8, 8)
        cr.fill()
        cr.set_source_rgb(1, 1, 1)
        cr.rectangle(5.5, 5.5, 8, 8)
        cr.stroke()
    elif name == "show":
        _rr(cr, 1.5, 2.5, 13, 9, 1)
        cr.stroke()
        cr.move_to(6.5, 4.8)
        cr.line_to(10.5, 7)
        cr.line_to(6.5, 9.2)
        cr.close_path()
        cr.fill()
        L((8, 11.5), (8, 14))
        L((5, 14.5), (11, 14.5))
    elif name == "notes":
        _rr(cr, 2, 1.5, 12, 13, 1)
        cr.stroke()
        L((4.5, 9), (11.5, 9))
        L((4.5, 11.5), (9.5, 11.5))
        cr.rectangle(2, 1.5, 12, 5)
        cr.stroke()
    elif name == "pdf":
        L((4, 1.5), (10, 1.5), (13, 4.5), (13, 14.5), (4, 14.5), (4, 1.5))
        L((10, 1.5), (10, 4.5), (13, 4.5))
    elif name == "bg":
        _rr(cr, 1.5, 2.5, 13, 11, 1)
        cr.fill()
    else:
        _orig_extra(cr, name, L, pi)


class PPTWindow(Gtk.ApplicationWindow):
    def __init__(self, app, path=None):
        super().__init__(application=app)
        self.app = app
        self.st = app.settings
        self.path = None
        self.pres = new_pres()
        self.cur = 0
        self.sel = []
        self.drag = None
        self.editing = None
        self.dirty = False
        self.undo_stack, self.redo_stack = [], []
        self.clip = None
        self.guides = []
        self.updating = False
        self.set_icon_name("pi-powerpoint")
        add_cls(self, "fx")
        self.set_default_size(self.st["width"], self.st["height"])
        if self.st.get("max"):
            self.maximize()
        self.connect("delete-event", self._on_close)
        self.connect("configure-event", self._cfg)
        self.connect("key-press-event", self._win_key)

        tb = WinTitlebar(self)
        t, self.title_lb = app_title("Presentation1 - PowerPoint", draw_logo)
        tb.pack_start(t)
        qs = Gtk.Box()
        qs.set_margin_start(12)
        qs.pack_start(icon_button("save", lambda w: self.save(), tip="Save (Ctrl+S)"), False, False, 0)
        qs.pack_start(icon_button("undo", lambda w: self.undo(), tip="Undo (Ctrl+Z)"), False, False, 0)
        qs.pack_start(icon_button("redo", lambda w: self.redo(), tip="Redo (Ctrl+Y)"), False, False, 0)
        qs.pack_start(icon_button("show", lambda w: self.slideshow(0), tip="Start from beginning (F5)"), False, False, 0)
        tb.pack_start(qs)
        tb.show_all()

        root = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.tabs = Gtk.Box()
        add_cls(self.tabs, "ribbon-tabs")
        self.ribbon = Gtk.Stack()
        self.tab_btns = {}
        fb = Gtk.Button(label="File")
        add_cls(fb, "rtab")
        fb.connect("clicked", self.file_menu)
        self.tabs.pack_start(fb, False, False, 0)
        for key, name, builder in (("home", "Home", self.rib_home), ("insert", "Insert", self.rib_insert),
                                   ("design", "Design", self.rib_design), ("show", "Slide Show", self.rib_show),
                                   ("view", "View", self.rib_view)):
            b = Gtk.Button(label=name)
            add_cls(b, "rtab")
            b.connect("clicked", lambda w, k=key: self.show_tab(k))
            self.tabs.pack_start(b, False, False, 0)
            self.tab_btns[key] = b
            box = Gtk.Box(spacing=2)
            add_cls(box, "ribbon")
            builder(box)
            self.ribbon.add_named(box, key)
        root.pack_start(self.tabs, False, False, 0)
        root.pack_start(self.ribbon, False, False, 0)

        body = Gtk.Paned(orientation=Gtk.Orientation.HORIZONTAL)
        tsw = Gtk.ScrolledWindow()
        tsw.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        add_cls(tsw, "thumbs")
        tsw.set_size_request(230, -1)
        self.thumbs = Gtk.ListBox()
        add_cls(self.thumbs, "thumbs")
        self.thumbs.set_selection_mode(Gtk.SelectionMode.BROWSE)
        self.thumbs.connect("row-selected", self._thumb_selected)
        self.thumbs.connect("key-press-event", self._thumb_key)
        self.thumbs.connect("button-press-event", self._thumb_press)
        tsw.add(self.thumbs)
        body.pack1(tsw, False, False)

        right = Gtk.Paned(orientation=Gtk.Orientation.VERTICAL)
        self.ov = Gtk.Overlay()
        add_cls(self.ov, "canvas-bg")
        self.da = Gtk.DrawingArea()
        self.da.set_can_focus(True)
        self.da.add_events(Gdk.EventMask.BUTTON_PRESS_MASK | Gdk.EventMask.BUTTON_RELEASE_MASK |
                           Gdk.EventMask.POINTER_MOTION_MASK | Gdk.EventMask.KEY_PRESS_MASK | Gdk.EventMask.SCROLL_MASK)
        self.da.connect("draw", self._draw)
        self.da.connect("button-press-event", self._press)
        self.da.connect("button-release-event", self._release)
        self.da.connect("motion-notify-event", self._motion)
        self.da.connect("key-press-event", self._canvas_key)
        self.da.connect("scroll-event", self._canvas_scroll)
        self.ov.add(self.da)
        self.editor = Gtk.TextView()
        self.editor.set_halign(Gtk.Align.START)
        self.editor.set_valign(Gtk.Align.START)
        self.editor.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
        self.editor.set_no_show_all(True)
        self.editor.get_buffer().connect("changed", self._ed_changed)
        self.editor.connect("key-press-event", self._ed_key)
        self.editor.connect("focus-out-event", lambda w, e: (GLib.idle_add(self._maybe_commit), False)[1])
        self.ov.add_overlay(self.editor)
        right.pack1(self.ov, True, False)
        nbox = Gtk.Box()
        add_cls(nbox, "notes")
        nsw = Gtk.ScrolledWindow()
        nsw.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        self.notes = Gtk.TextView()
        add_cls(self.notes, "notesv")
        self.notes.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
        self.notes.set_left_margin(16)
        self.notes.set_right_margin(16)
        self.notes.set_top_margin(8)
        self.notes.get_buffer().connect("changed", self._notes_changed)
        nsw.add(self.notes)
        self.notes_ov = Gtk.Overlay()
        self.notes_ov.add(nsw)
        self.notes_hint = label("Click to add notes", "ter")
        self.notes_hint.set_halign(Gtk.Align.START)
        self.notes_hint.set_valign(Gtk.Align.START)
        self.notes_hint.set_margin_start(18)
        self.notes_hint.set_margin_top(8)
        self.notes_ov.add_overlay(self.notes_hint)
        self.notes_ov.set_overlay_pass_through(self.notes_hint, True)
        nbox.pack_start(self.notes_ov, True, True, 0)
        nbox.set_size_request(-1, 90)
        self.notes_box = nbox
        right.pack2(nbox, False, True)
        self._paned_set = False

        def place_divider(w, a):
            if not self._paned_set and a.height > 300:
                self._paned_set = True
                w.set_position(a.height - 110)
        right.connect("size-allocate", place_divider)
        body.pack2(right, True, False)
        root.pack_start(body, True, True, 0)

        sb = Gtk.Box(spacing=16)
        add_cls(sb, "status")
        self.st_slide = label("", "sec")
        sb.pack_start(self.st_slide, False, False, 0)
        nb = Gtk.Button()
        add_cls(nb, "rb")
        nb.add(Gtk.Label(label="Notes"))
        nb.connect("clicked", lambda w: self.toggle_notes())
        sb.pack_end(nb, False, False, 0)
        sb.pack_end(icon_button("show", lambda w: self.slideshow(self.cur), tip="Slide show from current slide (Shift+F5)"),
                    False, False, 0)
        root.pack_start(sb, False, False, 0)
        self.add(root)
        root.show_all()
        self.notes_box.set_visible(self.st.get("notes", True))
        self.show_tab("home")
        self.load_pres(new_pres(), None)
        if path:
            GLib.idle_add(lambda: (self.open_path(path), False)[1])

    # ================================================================ helpers
    @property
    def slide(self):
        return self.pres["slides"][self.cur]

    @property
    def theme(self):
        return THEMES.get(self.pres["theme"], THEMES["office"])

    def geom(self):
        a = self.da.get_allocation()
        W, H = self.pres["w"], self.pres["h"]
        m = 28
        sc = max(0.05, min((a.width - 2 * m) / W, (a.height - 2 * m) / H))
        ox = (a.width - W * sc) / 2
        oy = (a.height - H * sc) / 2
        return sc, ox, oy

    def to_slide(self, x, y):
        sc, ox, oy = self.geom()
        return (x - ox) / sc, (y - oy) / sc

    def mark_dirty(self):
        if not self.dirty:
            self.dirty = True
            self.update_title()

    def changed(self, thumbs="cur"):
        self.mark_dirty()
        self.da.queue_draw()
        if thumbs == "all":
            for row in self.thumbs.get_children():
                row.get_child().get_children()[1].queue_draw()
        else:
            row = self.thumbs.get_row_at_index(self.cur)
            if row:
                row.get_child().get_children()[1].queue_draw()
        self._update_ribbon_state()

    # ================================================================ undo
    def push_undo(self):
        self.undo_stack.append((copy.deepcopy(self.pres), self.cur))
        del self.undo_stack[:-80]
        self.redo_stack.clear()

    def _restore(self, snap):
        self.pres, cur = snap
        self.sel = []
        self.build_thumbs()
        self.goto(min(cur, len(self.pres["slides"]) - 1))
        self.changed("all")

    def undo(self):
        if self.editing:
            self.commit_edit()
        if self.undo_stack:
            self.redo_stack.append((copy.deepcopy(self.pres), self.cur))
            self._restore(self.undo_stack.pop())

    def redo(self):
        if self.redo_stack:
            self.undo_stack.append((copy.deepcopy(self.pres), self.cur))
            self._restore(self.redo_stack.pop())

    # ================================================================ thumbnails
    def build_thumbs(self):
        for ch in self.thumbs.get_children():
            ch.destroy()
        W, H = self.pres["w"], self.pres["h"]
        tw = 176
        th = int(tw * H / W)
        for i, s in enumerate(self.pres["slides"]):
            row = Gtk.ListBoxRow()
            box = Gtk.Box()
            num = label(str(i + 1), "tnum")
            num.set_valign(Gtk.Align.START)
            box.pack_start(num, False, False, 0)
            da = Gtk.DrawingArea()
            da.set_size_request(tw, th)
            da.connect("draw", self._draw_thumb, row)
            box.pack_start(da, False, False, 0)
            row.add(box)
            self.thumbs.add(row)
        self.thumbs.show_all()

    def _draw_thumb(self, da, cr, row):
        i = row.get_index()
        if i < 0 or i >= len(self.pres["slides"]):
            return
        a = da.get_allocation()
        sc = a.width / self.pres["w"]
        sel = i == self.cur
        cr.save()
        cr.scale(sc, sc)
        render_slide(cr, self.pres, self.pres["slides"][i])
        cr.restore()
        cr.set_line_width(2 if sel else 1)
        cr.set_source_rgb(0.94, 0.44, 0.23) if sel else cr.set_source_rgba(1, 1, 1, 0.18)
        cr.rectangle(1 if sel else 0.5, 1 if sel else 0.5, a.width - (2 if sel else 1), a.height - (2 if sel else 1))
        cr.stroke()

    def _thumb_selected(self, lb, row):
        if row is None or self.updating:
            return
        if self.editing:
            self.commit_edit()
        self.goto(row.get_index(), from_list=True)

    def goto(self, i, from_list=False):
        i = max(0, min(i, len(self.pres["slides"]) - 1))
        self.cur = i
        self.sel = []
        if not from_list:
            self.updating = True
            row = self.thumbs.get_row_at_index(i)
            if row:
                self.thumbs.select_row(row)
                GLib.idle_add(self._scroll_thumb, row)
            self.updating = False
        self.updating = True
        self.notes.get_buffer().set_text(self.slide.get("notes", ""))
        self.updating = False
        self.notes_hint.set_visible(not self.slide.get("notes"))
        self.st_slide.set_text(f"Slide {i + 1} of {len(self.pres['slides'])}")
        for row in self.thumbs.get_children():
            row.get_child().get_children()[1].queue_draw()
        self._update_ribbon_state()
        self.da.queue_draw()

    def _scroll_thumb(self, row):
        adj = self.thumbs.get_parent().get_vadjustment() if isinstance(self.thumbs.get_parent(), Gtk.Viewport) else None
        if adj is None:
            return False
        a = row.get_allocation()
        if a.y < adj.get_value():
            adj.set_value(a.y)
        elif a.y + a.height > adj.get_value() + adj.get_page_size():
            adj.set_value(a.y + a.height - adj.get_page_size())
        return False

    def _thumb_key(self, w, ev):
        k = ev.keyval
        ctrl = ev.state & Gdk.ModifierType.CONTROL_MASK
        if k in (Gdk.KEY_Delete, Gdk.KEY_BackSpace):
            self.delete_slide()
            return True
        if ctrl and k in (Gdk.KEY_Up, Gdk.KEY_Down):
            self.move_slide(-1 if k == Gdk.KEY_Up else 1)
            return True
        if k == Gdk.KEY_Return:
            self.new_slide()
            return True
        return False

    def _thumb_press(self, w, ev):
        if ev.button != 3:
            return False
        row = self.thumbs.get_row_at_y(int(ev.y))
        if row is not None:
            self.thumbs.select_row(row)
        items = [{"label": "New slide", "icon": "newslide", "accel": "Ctrl+M", "cb": self.new_slide},
                 {"label": "Duplicate slide", "icon": "dup", "cb": self.duplicate_slide},
                 {"label": "Delete slide", "icon": "delete", "cb": self.delete_slide,
                  "sensitive": len(self.pres["slides"]) > 1}, None,
                 {"label": "Layout", "sub": [{"label": n, "cb": lambda k=k: self.set_layout(k)} for k, n in LAYOUTS]},
                 {"label": "Move up", "cb": lambda: self.move_slide(-1), "sensitive": self.cur > 0},
                 {"label": "Move down", "cb": lambda: self.move_slide(1), "sensitive": self.cur < len(self.pres["slides"]) - 1}]
        popup_menu(build_menu(items), self.thumbs, ev)
        return True

    # ================================================================ slides
    def new_slide(self, layout=None):
        if self.editing:
            self.commit_edit()
        self.push_undo()
        if layout is None:
            layout = "title_content" if self.cur > 0 or self.slide.get("layout") == "title" else self.slide.get("layout", "title_content")
        self.pres["slides"].insert(self.cur + 1, new_slide(layout, self.pres["w"], self.pres["h"]))
        self.build_thumbs()
        self.goto(self.cur + 1)
        self.changed()

    def duplicate_slide(self):
        self.push_undo()
        self.pres["slides"].insert(self.cur + 1, copy.deepcopy(self.slide))
        self.build_thumbs()
        self.goto(self.cur + 1)
        self.changed()

    def delete_slide(self):
        if len(self.pres["slides"]) < 2:
            return
        self.push_undo()
        del self.pres["slides"][self.cur]
        self.build_thumbs()
        self.goto(min(self.cur, len(self.pres["slides"]) - 1))
        self.changed()

    def move_slide(self, d):
        j = self.cur + d
        if not 0 <= j < len(self.pres["slides"]):
            return
        self.push_undo()
        s = self.pres["slides"]
        s[self.cur], s[j] = s[j], s[self.cur]
        self.build_thumbs()
        self.goto(j)
        self.changed("all")

    def set_layout(self, layout):
        """Swap the placeholders for the new layout, keeping text that was typed."""
        self.push_undo()
        sl = self.slide
        old = [s for s in sl["shapes"] if s["kind"] == "text" and s.get("role")]
        others = [s for s in sl["shapes"] if not (s["kind"] == "text" and s.get("role"))]
        new = layout_shapes(layout, self.pres["w"], self.pres["h"])
        texts = {"title": [], "subtitle": [], "body": []}
        for s in old:
            if s["text"]:
                texts[s["role"]].append(s["text"])
        for s in new:
            pool = texts[s["role"]] or (texts["subtitle"] if s["role"] == "body" else texts["body"] if s["role"] == "subtitle" else [])
            if pool:
                s["text"] = pool.pop(0)
        leftover = texts["title"] + texts["subtitle"] + texts["body"]
        for t in leftover:
            tb = text_shape(None, 80, 420, self.pres["w"] - 160, 60, 18)
            tb["text"] = t
            others.append(tb)
        sl["shapes"] = new + others
        sl["layout"] = layout
        self.sel = []
        self.changed()

    # ================================================================ drawing
    def _draw(self, w, cr):
        sc, ox, oy = self.geom()
        cr.set_source_rgb(0.106, 0.106, 0.106)
        cr.paint()
        W, H = self.pres["w"], self.pres["h"]
        cr.save()
        cr.translate(ox + 3, oy + 3)
        cr.set_source_rgba(0, 0, 0, 0.5)
        cr.rectangle(0, 0, W * sc, H * sc)
        cr.fill()
        cr.restore()
        cr.save()
        cr.translate(ox, oy)
        cr.scale(sc, sc)
        render_slide(cr, self.pres, self.slide, edit=True, hide=self.editing)
        cr.restore()
        # selection
        for s in self.sel:
            self._draw_sel(cr, s, sc, ox, oy, single=len(self.sel) == 1)
        if self.drag and self.drag[0] == "marquee":
            _, x0, y0, x1, y1 = self.drag
            cr.set_source_rgba(0.94, 0.44, 0.23, 0.15)
            cr.rectangle(min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0))
            cr.fill_preserve()
            cr.set_source_rgb(0.94, 0.44, 0.23)
            cr.set_line_width(1)
            cr.stroke()
        cr.set_source_rgba(0.94, 0.44, 0.23, 0.9)
        cr.set_line_width(1)
        cr.set_dash([5, 4])
        for kind, v in self.guides:
            if kind == "v":
                x = ox + v * sc
                cr.move_to(int(x) + 0.5, oy)
                cr.line_to(int(x) + 0.5, oy + H * sc)
            else:
                y = oy + v * sc
                cr.move_to(ox, int(y) + 0.5)
                cr.line_to(ox + W * sc, int(y) + 0.5)
        cr.stroke()
        cr.set_dash([])

    def _draw_sel(self, cr, s, sc, ox, oy, single):
        x, y, w, h = ox + s["x"] * sc, oy + s["y"] * sc, s["w"] * sc, s["h"] * sc
        cr.set_line_width(1)
        if s["kind"] == "line":
            pts = [(x, y), (x + w, y + h)]
        else:
            nx, ny, nw, nh = min(x, x + w), min(y, y + h), abs(w), abs(h)
            cr.set_source_rgb(0.55, 0.55, 0.55)
            if s is self.editing:
                cr.set_dash([4, 3])
            cr.rectangle(int(nx) + 0.5, int(ny) + 0.5, int(nw), int(nh))
            cr.stroke()
            cr.set_dash([])
            pts = self.handles(s, sc, ox, oy)
        if not single or s is self.editing:
            return
        for px, py in pts:
            cr.arc(px, py, 4.5, 0, 2 * math.pi)
            cr.set_source_rgb(1, 1, 1)
            cr.fill_preserve()
            cr.set_source_rgb(0.4, 0.4, 0.4)
            cr.stroke()

    def handles(self, s, sc, ox, oy):
        x, y, w, h = ox + s["x"] * sc, oy + s["y"] * sc, s["w"] * sc, s["h"] * sc
        if s["kind"] == "line":
            return [(x, y), (x + w, y + h)]
        return [(x, y), (x + w / 2, y), (x + w, y), (x + w, y + h / 2), (x + w, y + h), (x + w / 2, y + h), (x, y + h),
                (x, y + h / 2)]

    # ================================================================ mouse
    def hit(self, px, py):
        sc, ox, oy = self.geom()
        tol = 5 / sc
        for s in reversed(self.slide["shapes"]):
            if s["kind"] == "line":
                x1, y1, x2, y2 = s["x"], s["y"], s["x"] + s["w"], s["y"] + s["h"]
                dx, dy = x2 - x1, y2 - y1
                L2 = dx * dx + dy * dy or 1
                t = max(0, min(1, ((px - x1) * dx + (py - y1) * dy) / L2))
                if math.hypot(px - (x1 + t * dx), py - (y1 + t * dy)) <= tol + s.get("line_w", 2) / 2:
                    return s
                continue
            x, y = min(s["x"], s["x"] + s["w"]), min(s["y"], s["y"] + s["h"])
            if x - tol <= px <= x + abs(s["w"]) + tol and y - tol <= py <= y + abs(s["h"]) + tol:
                return s
        return None

    def _press(self, w, ev):
        sc, ox, oy = self.geom()
        px, py = self.to_slide(ev.x, ev.y)
        if self.editing:
            ed = self.editing
            if ev.button == 1 and ed["x"] <= px <= ed["x"] + ed["w"] and ed["y"] <= py <= ed["y"] + ed["h"]:
                return False
            self.commit_edit()
        self.da.grab_focus()
        if ev.type == Gdk.EventType._2BUTTON_PRESS and ev.button == 1:
            s = self.hit(px, py)
            if s is not None:
                if s["kind"] == "table":
                    self.edit_table(s)
                elif s["kind"] not in ("image", "line"):
                    self.start_edit(s)
            return True
        if ev.button == 3:
            s = self.hit(px, py)
            if s is not None and s not in self.sel:
                self.sel = [s]
                self.da.queue_draw()
            self.shape_menu(ev)
            return True
        if ev.button != 1:
            return False
        # handles of the single selected shape
        if len(self.sel) == 1:
            s = self.sel[0]
            for i, (hx, hy) in enumerate(self.handles(s, sc, ox, oy)):
                if abs(ev.x - hx) <= 6 and abs(ev.y - hy) <= 6:
                    self.push_undo()
                    self.drag = ("resize", s, i, px, py, dict(s))
                    return True
        s = self.hit(px, py)
        shift = ev.state & (Gdk.ModifierType.SHIFT_MASK | Gdk.ModifierType.CONTROL_MASK)
        if s is None:
            if not shift:
                self.sel = []
            self.drag = ("marquee", ev.x, ev.y, ev.x, ev.y)
            self._update_ribbon_state()
            self.da.queue_draw()
            return True
        if shift:
            if s in self.sel:
                self.sel.remove(s)
            else:
                self.sel.append(s)
            self.da.queue_draw()
            self._update_ribbon_state()
            return True
        was_selected = s in self.sel and len(self.sel) == 1
        if s not in self.sel:
            self.sel = [s]
        # an empty placeholder goes straight into typing, like PowerPoint
        if s["kind"] == "text" and s.get("role") and not s["text"]:
            self.start_edit(s)
            return True
        self.push_undo()
        self.drag = ("move", px, py, [(x, dict(x)) for x in self.sel], was_selected, False)
        self._update_ribbon_state()
        self.da.queue_draw()
        return True

    def _motion(self, w, ev):
        sc, ox, oy = self.geom()
        px, py = self.to_slide(ev.x, ev.y)
        win = self.da.get_window()
        if self.drag is None:
            name = None
            if len(self.sel) == 1 and not self.editing:
                for i, (hx, hy) in enumerate(self.handles(self.sel[0], sc, ox, oy)):
                    if abs(ev.x - hx) <= 6 and abs(ev.y - hy) <= 6:
                        name = "move" if self.sel[0]["kind"] == "line" else \
                            ["nwse-resize", "ns-resize", "nesw-resize", "ew-resize"][i % 4]
                        break
            if name is None:
                s = self.hit(px, py)
                if s is not None:
                    name = "text" if s is self.editing else "move"
            win.set_cursor(Gdk.Cursor.new_from_name(self.da.get_display(), name) if name else None)
            return False
        k = self.drag[0]
        W, H = self.pres["w"], self.pres["h"]
        if k == "move":
            _, x0, y0, orig, was, moved = self.drag
            dx, dy = px - x0, py - y0
            if not moved and math.hypot(dx * sc, dy * sc) < 3:
                return True
            self.drag = ("move", x0, y0, orig, was, True)
            if ev.state & Gdk.ModifierType.SHIFT_MASK:
                if abs(dx) > abs(dy):
                    dy = 0
                else:
                    dx = 0
            # snap the group to the slide center and edges
            bx0 = min(min(o["x"], o["x"] + o["w"]) for _, o in orig) + dx
            by0 = min(min(o["y"], o["y"] + o["h"]) for _, o in orig) + dy
            bx1 = max(max(o["x"], o["x"] + o["w"]) for _, o in orig) + dx
            by1 = max(max(o["y"], o["y"] + o["h"]) for _, o in orig) + dy
            self.guides = []
            tol = 6 / sc
            for edge, tgt in (((bx0 + bx1) / 2, W / 2), (bx0, 0), (bx1, W)):
                if abs(edge - tgt) < tol:
                    dx += tgt - edge
                    self.guides.append(("v", tgt))
                    break
            for edge, tgt in (((by0 + by1) / 2, H / 2), (by0, 0), (by1, H)):
                if abs(edge - tgt) < tol:
                    dy += tgt - edge
                    self.guides.append(("h", tgt))
                    break
            for s, o in orig:
                s["x"], s["y"] = o["x"] + dx, o["y"] + dy
            self.changed()
        elif k == "resize":
            _, s, i, x0, y0, o = self.drag
            dx, dy = px - x0, py - y0
            if s["kind"] == "line":
                if i == 0:
                    s["x"], s["y"] = o["x"] + dx, o["y"] + dy
                    s["w"], s["h"] = o["w"] - dx, o["h"] - dy
                else:
                    s["w"], s["h"] = o["w"] + dx, o["h"] + dy
                if ev.state & Gdk.ModifierType.SHIFT_MASK:
                    if abs(s["w"]) > abs(s["h"]) * 2:
                        s["h"] = 0
                    elif abs(s["h"]) > abs(s["w"]) * 2:
                        s["w"] = 0
                self.changed()
                return True
            x, y, w_, h_ = o["x"], o["y"], o["w"], o["h"]
            left, top, right, bottom = x, y, x + w_, y + h_
            if i in (0, 6, 7):
                left += dx
            if i in (2, 3, 4):
                right += dx
            if i in (0, 1, 2):
                top += dy
            if i in (4, 5, 6):
                bottom += dy
            keep = (s["kind"] == "image") != bool(ev.state & Gdk.ModifierType.SHIFT_MASK)
            if keep and i in (0, 2, 4, 6) and w_ and h_:
                ratio = w_ / h_
                nw, nh = right - left, bottom - top
                if abs(nw / ratio) > abs(nh):
                    nh = nw / ratio
                else:
                    nw = nh * ratio
                if i in (0, 6):
                    left = right - nw
                else:
                    right = left + nw
                if i in (0, 2):
                    top = bottom - nh
                else:
                    bottom = top + nh
            s["x"], s["y"] = min(left, right), min(top, bottom)
            s["w"], s["h"] = max(4, abs(right - left)), max(4, abs(bottom - top))
            self.changed()
        elif k == "marquee":
            self.drag = ("marquee", self.drag[1], self.drag[2], ev.x, ev.y)
            self.da.queue_draw()
        elif k == "draw":
            _, kind, x0, y0, s = self.drag
            if ev.state & Gdk.ModifierType.SHIFT_MASK and kind != "line":
                d = max(abs(px - x0), abs(py - y0))
                px = x0 + math.copysign(d, px - x0)
                py = y0 + math.copysign(d, py - y0)
            if kind == "line":
                s["w"], s["h"] = px - x0, py - y0
            else:
                s["x"], s["y"] = min(x0, px), min(y0, py)
                s["w"], s["h"] = max(4, abs(px - x0)), max(4, abs(py - y0))
            self.changed()
        return True

    def _release(self, w, ev):
        d = self.drag
        self.drag = None
        self.guides = []
        if d is None:
            return False
        if d[0] == "move":
            if not d[5]:
                self.undo_stack.pop()
                if d[4] and len(self.sel) == 1 and self.sel[0]["kind"] not in ("image", "line", "table"):
                    # second click on a selected shape: type into it
                    self.start_edit(self.sel[0], click=(ev.x, ev.y))
        elif d[0] == "marquee":
            sc, ox, oy = self.geom()
            x0, y0 = self.to_slide(min(d[1], d[3]), min(d[2], d[4]))
            x1, y1 = self.to_slide(max(d[1], d[3]), max(d[2], d[4]))
            if x1 - x0 > 2 and y1 - y0 > 2:
                self.sel = [s for s in self.slide["shapes"]
                            if min(s["x"], s["x"] + s["w"]) >= x0 and max(s["x"], s["x"] + s["w"]) <= x1 and
                            min(s["y"], s["y"] + s["h"]) >= y0 and max(s["y"], s["y"] + s["h"]) <= y1]
            self._update_ribbon_state()
        elif d[0] == "draw":
            s = d[4]
            if abs(s["w"]) < 6 and abs(s["h"]) < 6:
                s["w"], s["h"] = (160, 0) if s["kind"] == "line" else (160, 100)
            self.sel = [s]
            self.changed()
            if s["kind"] != "line":
                self.start_edit(s) if s["kind"] == "text" else None
        self.da.queue_draw()
        return True

    def _canvas_scroll(self, w, ev):
        ok, dx, dy = ev.get_scroll_deltas()
        if not ok:
            dy = {Gdk.ScrollDirection.UP: -1, Gdk.ScrollDirection.DOWN: 1}.get(ev.direction, 0)
        if dy > 0 and self.cur < len(self.pres["slides"]) - 1:
            self.goto(self.cur + 1)
        elif dy < 0 and self.cur > 0:
            self.goto(self.cur - 1)
        return True

    # ================================================================ keyboard
    def _canvas_key(self, w, ev):
        k = ev.keyval
        ctrl = ev.state & Gdk.ModifierType.CONTROL_MASK
        shift = ev.state & Gdk.ModifierType.SHIFT_MASK
        if k in (Gdk.KEY_Delete, Gdk.KEY_BackSpace) and self.sel:
            self.push_undo()
            for s in self.sel:
                if s in self.slide["shapes"]:
                    if s["kind"] == "text" and s.get("role") and s["text"]:
                        s["text"] = ""          # placeholders empty out instead of vanishing
                    else:
                        self.slide["shapes"].remove(s)
            self.sel = []
            self.changed()
            return True
        if k in (Gdk.KEY_Delete,) and not self.sel:
            self.delete_slide()
            return True
        arrows = {Gdk.KEY_Left: (-1, 0), Gdk.KEY_Right: (1, 0), Gdk.KEY_Up: (0, -1), Gdk.KEY_Down: (0, 1)}
        if k in arrows and self.sel:
            dx, dy = arrows[k]
            step = 1 if ctrl else 10 if shift else 4
            self.push_undo()
            for s in self.sel:
                s["x"] += dx * step
                s["y"] += dy * step
            self.changed()
            return True
        if k in (Gdk.KEY_Page_Down, Gdk.KEY_Down, Gdk.KEY_Right) and not self.sel:
            self.goto(self.cur + 1)
            return True
        if k in (Gdk.KEY_Page_Up, Gdk.KEY_Up, Gdk.KEY_Left) and not self.sel:
            self.goto(self.cur - 1)
            return True
        if k == Gdk.KEY_Escape:
            self.sel = []
            self._update_ribbon_state()
            self.da.queue_draw()
            return True
        if k == Gdk.KEY_Tab:
            shapes = self.slide["shapes"]
            if shapes:
                i = shapes.index(self.sel[0]) if self.sel and self.sel[0] in shapes else -1
                self.sel = [shapes[(i + (-1 if shift else 1)) % len(shapes)]]
                self._update_ribbon_state()
                self.da.queue_draw()
            return True
        if k in (Gdk.KEY_Return, Gdk.KEY_F2) and len(self.sel) == 1 and self.sel[0]["kind"] not in ("image", "line"):
            if self.sel[0]["kind"] == "table":
                self.edit_table(self.sel[0])
            else:
                self.start_edit(self.sel[0])
            return True
        if ctrl:
            low = Gdk.keyval_to_lower(k)
            acts = {Gdk.KEY_c: self.copy_shapes, Gdk.KEY_x: lambda: self.copy_shapes(cut=True), Gdk.KEY_v: self.paste,
                    Gdk.KEY_d: self.duplicate_shapes, Gdk.KEY_a: self.select_all,
                    Gdk.KEY_b: lambda: self.toggle("b"), Gdk.KEY_i: lambda: self.toggle("i"), Gdk.KEY_u: lambda: self.toggle("u"),
                    Gdk.KEY_l: lambda: self.set_prop("align", "left"), Gdk.KEY_e: lambda: self.set_prop("align", "center"),
                    Gdk.KEY_r: lambda: self.set_prop("align", "right")}
            if low in acts:
                acts[low]()
                return True
            return False
        ch = Gdk.keyval_to_unicode(k)
        if ch and ch >= 32 and len(self.sel) == 1 and self.sel[0]["kind"] not in ("image", "line", "table"):
            self.start_edit(self.sel[0], initial=chr(ch))
            return True
        return False

    def _win_key(self, w, ev):
        k = ev.keyval
        ctrl = ev.state & Gdk.ModifierType.CONTROL_MASK
        shift = ev.state & Gdk.ModifierType.SHIFT_MASK
        if k == Gdk.KEY_F5:
            self.slideshow(self.cur if shift else 0)
            return True
        if not ctrl:
            return False
        low = Gdk.keyval_to_lower(k)
        acts = {Gdk.KEY_s: (self.save_as if shift else self.save), Gdk.KEY_o: self.open_dialog, Gdk.KEY_n: self.app.new_window,
                Gdk.KEY_w: self.close, Gdk.KEY_m: self.new_slide, Gdk.KEY_z: self.undo, Gdk.KEY_y: self.redo,
                Gdk.KEY_p: self.export_pdf}
        if self.notes.has_focus() and low in (Gdk.KEY_z, Gdk.KEY_y):
            return False
        if low in acts:
            if self.editing and low not in (Gdk.KEY_z, Gdk.KEY_y):
                self.commit_edit()
            acts[low]()
            return True
        return False

    # ================================================================ text editing
    def start_edit(self, s, initial=None, click=None):
        if self.editing:
            self.commit_edit()
        self.push_undo()
        self.editing = s
        self.sel = [s]
        buf = self.editor.get_buffer()
        self.updating = True
        buf.set_text(s["text"] + (initial or ""))
        self.updating = False
        self._style_editor()
        self.editor.show()
        self.editor.grab_focus()
        if click is not None:
            sc, ox, oy = self.geom()
            bx, by = self.editor.window_to_buffer_coords(Gtk.TextWindowType.TEXT,
                                                        int(click[0] - self.editor.get_margin_start()),
                                                        int(click[1] - self.editor.get_margin_top()))
            ok, it = self.editor.get_iter_at_location(bx, by)
            buf.place_cursor(it if ok else buf.get_end_iter())
        else:
            buf.place_cursor(buf.get_end_iter())
        self._update_ribbon_state()
        self.da.queue_draw()

    def _style_editor(self):
        s = self.editing
        if s is None:
            return
        sc, ox, oy = self.geom()
        th = self.theme
        role = s.get("role")
        fam = font_for(s.get("font") or (th["tfont"] if role == "title" else th["bfont"]))
        d = Pango.FontDescription()
        d.set_family(fam)
        d.set_absolute_size(s["size"] * sc * Pango.SCALE)
        d.set_weight(Pango.Weight.BOLD if s.get("b") else Pango.Weight.NORMAL)
        d.set_style(Pango.Style.ITALIC if s.get("i") else Pango.Style.NORMAL)
        self.editor.override_font(d)
        if s["kind"] == "text":
            col = s.get("color") or (th["title"] if role == "title" else th["text"])
            bg = s.get("fill") or self.slide.get("bg") or th["bg"]
        else:
            bg = s.get("fill", th["accent"]) or self.slide.get("bg") or th["bg"]
            col = s.get("color") or ("#FFFFFF" if luminance(bg) < 0.6 else "#000000")
        rgba = Gdk.RGBA()
        rgba.parse(col)
        self.editor.override_color(Gtk.StateFlags.NORMAL, rgba)
        bga = Gdk.RGBA()
        bga.parse(bg)
        self.editor.override_background_color(Gtk.StateFlags.NORMAL, bga)
        sel_bg = Gdk.RGBA()
        sel_bg.parse("#F4B183")
        self.editor.override_background_color(Gtk.StateFlags.SELECTED, sel_bg)
        al = s.get("align", "center" if s["kind"] != "text" else "left")
        self.editor.set_justification({"left": Gtk.Justification.LEFT, "center": Gtk.Justification.CENTER,
                                       "right": Gtk.Justification.RIGHT}[al])
        tabs = Pango.TabArray.new(1, True)
        tabs.set_tab(0, Pango.TabAlign.LEFT, int(36 * sc))
        self.editor.set_tabs(tabs)
        x, y = ox + min(s["x"], s["x"] + s["w"]) * sc, oy + min(s["y"], s["y"] + s["h"]) * sc
        w, h = abs(s["w"]) * sc, abs(s["h"]) * sc
        ind = s["size"] * 0.9 * sc if s.get("bullets") else 0
        self.editor.set_left_margin(int(7.2 * sc + ind))
        self.editor.set_right_margin(int(7.2 * sc))
        self.editor.set_margin_start(max(0, int(x)))
        self.editor.set_margin_top(max(0, int(y)))
        self.editor.set_size_request(max(20, int(w)), max(10, int(h)))
        self._valign_editor()

    def _valign_editor(self):
        s = self.editing
        if s is None:
            return
        sc, ox, oy = self.geom()
        h = abs(s["h"]) * sc
        va = s.get("valign", "middle" if s["kind"] != "text" else "top")
        buf = self.editor.get_buffer()
        text = buf.get_text(buf.get_start_iter(), buf.get_end_iter(), False)
        surf = cairo.ImageSurface(cairo.FORMAT_ARGB32, 1, 1)
        cr = cairo.Context(surf)
        cr.scale(sc, sc)
        th = text_height(para_layouts(cr, dict(s, text=text or " "), self.theme)) * sc
        pad = 3.6 * sc
        top = pad if va == "top" else max(pad, (h - th) / 2) if va == "middle" else max(pad, h - th - pad)
        self.editor.set_top_margin(int(top))

    def _ed_changed(self, buf):
        if self.updating or self.editing is None:
            return
        self.editing["text"] = buf.get_text(buf.get_start_iter(), buf.get_end_iter(), False)
        self._valign_editor()
        self.mark_dirty()
        row = self.thumbs.get_row_at_index(self.cur)
        if row:
            row.get_child().get_children()[1].queue_draw()

    def _ed_key(self, w, ev):
        if ev.keyval == Gdk.KEY_Escape:
            s = self.editing
            self.commit_edit()
            self.sel = [s] if s and s in self.slide["shapes"] else []
            self.da.grab_focus()
            self.da.queue_draw()
            return True
        if ev.keyval in (Gdk.KEY_Tab, Gdk.KEY_ISO_Left_Tab) and self.editing and self.editing.get("bullets"):
            buf = self.editor.get_buffer()
            it = buf.get_iter_at_mark(buf.get_insert())
            ls = it.copy()
            ls.set_line_offset(0)
            if ev.keyval == Gdk.KEY_ISO_Left_Tab or ev.state & Gdk.ModifierType.SHIFT_MASK:
                if ls.get_char() == "\t":
                    e = ls.copy()
                    e.forward_char()
                    buf.delete(ls, e)
            else:
                buf.insert(ls, "\t")
            return True
        ctrl = ev.state & Gdk.ModifierType.CONTROL_MASK
        if ctrl:
            low = Gdk.keyval_to_lower(ev.keyval)
            if low in (Gdk.KEY_b, Gdk.KEY_i, Gdk.KEY_u):
                self.toggle({Gdk.KEY_b: "b", Gdk.KEY_i: "i", Gdk.KEY_u: "u"}[low])
                return True
            if low in (Gdk.KEY_l, Gdk.KEY_e, Gdk.KEY_r):
                self.set_prop("align", {Gdk.KEY_l: "left", Gdk.KEY_e: "center", Gdk.KEY_r: "right"}[low])
                return True
            if low == Gdk.KEY_z:
                buf = self.editor.get_buffer()
                if hasattr(buf, "undo"):
                    return False
        return False

    def _maybe_commit(self):
        if self.editing and not self.editor.has_focus():
            # keep editing while the ribbon (font box, color menu...) is in use
            fw = self.get_focus()
            if fw is not None and fw.is_ancestor(self.ribbon):
                return False
            self.commit_edit()
        return False

    def commit_edit(self):
        s = self.editing
        if s is None:
            return
        self.editing = None
        self.editor.hide()
        buf = self.editor.get_buffer()
        s["text"] = buf.get_text(buf.get_start_iter(), buf.get_end_iter(), False).rstrip("\n") if s["kind"] == "text" \
            else buf.get_text(buf.get_start_iter(), buf.get_end_iter(), False)
        if self.undo_stack:
            prev = self.undo_stack[-1][0]["slides"]
            same = self.cur < len(prev) and prev[self.cur] == self.slide
            if same:
                self.undo_stack.pop()
        # a plain text box left empty disappears
        if s["kind"] == "text" and not s.get("role") and not s["text"].strip() and not s.get("fill") and s in self.slide["shapes"]:
            self.slide["shapes"].remove(s)
            self.sel = []
        self.changed()

    # ================================================================ shape operations
    def select_all(self):
        self.sel = list(self.slide["shapes"])
        self._update_ribbon_state()
        self.da.queue_draw()

    def copy_shapes(self, cut=False):
        if not self.sel:
            return
        self.clip = [copy.deepcopy(s) for s in self.sel]
        if cut:
            self.push_undo()
            for s in self.sel:
                if s in self.slide["shapes"]:
                    self.slide["shapes"].remove(s)
            self.sel = []
            self.changed()

    def paste(self):
        if not self.clip:
            # a copied picture from another app
            cb = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)
            pb = cb.wait_for_image()
            if pb is not None:
                ok, data = pb.save_to_bufferv("png", [], [])
                self.add_picture(bytes(data))
                return
            text = cb.wait_for_text()
            if text:
                self.push_undo()
                tb = text_shape(None, 120, 200, self.pres["w"] - 240, 60, 18)
                tb["text"] = text
                self.slide["shapes"].append(tb)
                self.sel = [tb]
                self.changed()
            return
        self.push_undo()
        new = []
        for s in self.clip:
            n = copy.deepcopy(s)
            if any(abs(o["x"] - n["x"]) < 1 and abs(o["y"] - n["y"]) < 1 for o in self.slide["shapes"]):
                n["x"] += 18
                n["y"] += 18
            n["role"] = None if n.get("role") and any(o.get("role") == n["role"] for o in self.slide["shapes"]) else n.get("role")
            self.slide["shapes"].append(n)
            new.append(n)
        self.clip = [copy.deepcopy(s) for s in new]
        self.sel = new
        self.changed()

    def duplicate_shapes(self):
        if not self.sel:
            self.duplicate_slide()
            return
        self.push_undo()
        new = []
        for s in self.sel:
            n = copy.deepcopy(s)
            n["x"] += 18
            n["y"] += 18
            n["role"] = None
            self.slide["shapes"].append(n)
            new.append(n)
        self.sel = new
        self.changed()

    def arrange(self, how):
        if not self.sel:
            return
        self.push_undo()
        sh = self.slide["shapes"]
        for s in (self.sel if how in ("front", "forward") else reversed(self.sel)):
            i = sh.index(s)
            sh.pop(i)
            if how == "front":
                sh.append(s)
            elif how == "back":
                sh.insert(0, s)
            elif how == "forward":
                sh.insert(min(len(sh), i + 1), s)
            else:
                sh.insert(max(0, i - 1), s)
        self.changed()

    def targets(self):
        if self.editing:
            return [self.editing]
        return [s for s in self.sel if s["kind"] not in ("image",)]

    def set_prop(self, key, value):
        t = self.targets()
        if not t:
            return
        if not self.editing:
            self.push_undo()
        for s in t:
            s[key] = value
        if self.editing:
            self._style_editor()
            self.editor.grab_focus()
        self.changed()

    def toggle(self, key):
        t = self.targets()
        if not t:
            return
        self.set_prop(key, not all(s.get(key) for s in t))

    def grow_font(self, d):
        t = self.targets()
        if not t:
            return
        if not self.editing:
            self.push_undo()
        for s in t:
            cur = s.get("size", 18)
            bigger = [z for z in SIZES if z > cur]
            smaller = [z for z in SIZES if z < cur]
            s["size"] = (bigger[0] if bigger else cur + 8) if d > 0 else (smaller[-1] if smaller else max(6, cur - 2))
        if self.editing:
            self._style_editor()
        self.changed()

    def insert_textbox(self):
        self._start_draw("text")

    def _start_draw(self, kind):
        if self.editing:
            self.commit_edit()
        W, H = self.pres["w"], self.pres["h"]
        self.push_undo()
        if kind == "text":
            s = text_shape(None, W / 2 - 200, H / 2 - 30, 400, 60, 18)
        elif kind == "line":
            s = {"kind": "line", "x": W / 2 - 120, "y": H / 2, "w": 240, "h": 0, "line": None, "line_w": 2.25}
        else:
            s = {"kind": kind, "x": W / 2 - 100, "y": H / 2 - 70, "w": 200, "h": 140, "text": "", "font": None, "size": 18,
                 "b": False, "i": False, "u": False, "color": None, "align": "center", "valign": "middle", "bullets": False,
                 "fill": self.theme["accent"], "line": darker(self.theme["accent"])}
            if kind in ("ellipse", "star"):
                s["w"] = s["h"] = 140
                s["x"] = W / 2 - 70
        self.slide["shapes"].append(s)
        self.sel = [s]
        self.changed()
        if kind == "text":
            self.start_edit(s)

    def insert_picture(self):
        dlg = Gtk.FileChooserNative.new("Insert picture", self, Gtk.FileChooserAction.OPEN, "Insert", "Cancel")
        f = Gtk.FileFilter()
        f.set_name("Pictures")
        f.add_pixbuf_formats()
        dlg.add_filter(f)
        dlg.set_current_folder(os.path.join(HOME, "Pictures"))
        if dlg.run() == Gtk.ResponseType.ACCEPT:
            try:
                with open(dlg.get_filename(), "rb") as fh:
                    data = fh.read()
                if not data.startswith((b"\x89PNG", b"\xff\xd8", b"GIF8")):
                    pb = GdkPixbuf.Pixbuf.new_from_file(dlg.get_filename())
                    ok, b = pb.save_to_bufferv("png", [], [])
                    data = bytes(b)
                self.add_picture(data)
            except (OSError, GLib.Error) as e:
                info_dialog(self, "Couldn't insert the picture", str(e))

    def add_picture(self, data):
        pb = pixbuf_for(data)
        if pb is None:
            return
        W, H = self.pres["w"], self.pres["h"]
        w, h = pb.get_width() * 0.75, pb.get_height() * 0.75
        k = min(1.0, W * 0.6 / w, H * 0.7 / h)
        w, h = w * k, h * k
        self.push_undo()
        s = {"kind": "image", "x": (W - w) / 2, "y": (H - h) / 2, "w": w, "h": h, "img": data}
        self.slide["shapes"].append(s)
        self.sel = [s]
        self.changed()

    def insert_table(self):
        g = Gtk.Grid()
        g.set_column_spacing(10)
        g.set_row_spacing(8)
        cols = Gtk.SpinButton.new_with_range(1, 10, 1)
        cols.set_value(3)
        rows = Gtk.SpinButton.new_with_range(1, 20, 1)
        rows.set_value(4)
        g.attach(label("Columns"), 0, 0, 1, 1)
        g.attach(cols, 1, 0, 1, 1)
        g.attach(label("Rows"), 0, 1, 1, 1)
        g.attach(rows, 1, 1, 1, 1)
        dlg = WinDialog(self, "Insert table", primary="Insert", secondary="Cancel", width=320, extra=g)

        def go(ok):
            if not ok:
                return
            nc, nr = int(cols.get_value()), int(rows.get_value())
            W, H = self.pres["w"], self.pres["h"]
            w = min(W - 120, nc * 180)
            h = nr * 40
            self.push_undo()
            s = {"kind": "table", "x": (W - w) / 2, "y": max(130, (H - h) / 2), "w": w, "h": h,
                 "rows": [[f"Header {j + 1}" if i == 0 else "" for j in range(nc)] for i in range(nr)], "size": 16}
            self.slide["shapes"].append(s)
            self.sel = [s]
            self.changed()
            self.edit_table(s)
        dlg.ask(go)

    def edit_table(self, s):
        rows = s["rows"]
        nr, nc = len(rows), max(len(r) for r in rows)
        grid = Gtk.Grid()
        grid.set_row_spacing(4)
        grid.set_column_spacing(4)
        ents = []
        state = {"nr": nr, "nc": nc}

        def build():
            for ch in grid.get_children():
                ch.destroy()
            old = [[e.get_text() for e in r] for r in ents] if ents else rows
            ents.clear()
            for i in range(state["nr"]):
                row = []
                for j in range(state["nc"]):
                    e = Gtk.Entry()
                    e.set_width_chars(12)
                    e.set_text(old[i][j] if i < len(old) and j < len(old[i]) else "")
                    grid.attach(e, j, i, 1, 1)
                    row.append(e)
                ents.append(row)
            grid.show_all()
        build()
        bx = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        sw = Gtk.ScrolledWindow()
        sw.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        sw.set_min_content_height(min(320, 40 * nr + 10))
        sw.set_min_content_width(min(720, 130 * nc))
        sw.add(grid)
        bx.pack_start(sw, True, True, 0)
        btns = Gtk.Box(spacing=6)
        for text, dr, dc in (("Add row", 1, 0), ("Add column", 0, 1), ("Remove row", -1, 0), ("Remove column", 0, -1)):
            b = Gtk.Button(label=text)
            add_cls(b, "secondary")

            def go(w, dr=dr, dc=dc):
                state["nr"] = max(1, state["nr"] + dr)
                state["nc"] = max(1, state["nc"] + dc)
                build()
            b.connect("clicked", go)
            btns.pack_start(b, False, False, 0)
        bx.pack_start(btns, False, False, 0)
        dlg = WinDialog(self, "Edit table", primary="Done", secondary="Cancel", width=760, extra=bx)

        def done(ok):
            if not ok:
                return
            self.push_undo()
            new = [[e.get_text() for e in r] for r in ents]
            if len(new) != len(rows):
                s["h"] = s["h"] / max(1, len(rows)) * len(new)
            s["rows"] = new
            self.changed()
        dlg.ask(done)

    # ================================================================ themes / design
    def set_theme(self, key):
        self.push_undo()
        self.pres["theme"] = key
        for sl in self.pres["slides"]:
            sl["bg"] = None
            for s in sl["shapes"]:
                if s["kind"] not in ("text", "image", "table", "line") and s.get("fill") in [t["accent"] for t in THEMES.values()]:
                    s["fill"] = THEMES[key]["accent"]
                    s["line"] = darker(THEMES[key]["accent"])
        for k, b in self.theme_btns.items():
            (b.get_style_context().add_class if k == key else b.get_style_context().remove_class)("on")
        self.changed("all")

    def set_bg(self, color, all_slides=False):
        self.push_undo()
        for sl in (self.pres["slides"] if all_slides else [self.slide]):
            sl["bg"] = color
        self.changed("all" if all_slides else "cur")

    def set_size(self, wide):
        W = 960.0 if wide else 720.0
        if abs(W - self.pres["w"]) < 1:
            return
        self.push_undo()
        k = W / self.pres["w"]
        for sl in self.pres["slides"]:
            for s in sl["shapes"]:
                s["x"] *= k
                s["w"] *= k
        self.pres["w"] = W
        self.build_thumbs()
        self.goto(self.cur)
        self.changed("all")

    # ================================================================ ribbon
    def show_tab(self, key):
        for k, b in self.tab_btns.items():
            (b.get_style_context().add_class if k == key else b.get_style_context().remove_class)("sel")
        self.ribbon.set_visible_child_name(key)

    def rsep(self, box):
        s = Gtk.Box()
        add_cls(s, "rsep")
        box.pack_start(s, False, False, 0)

    def rbtn(self, box, icon, tip, cb, text=None, markup=None):
        b = Gtk.Button()
        add_cls(b, "rb")
        b.set_can_focus(False)
        h = Gtk.Box(spacing=6)
        if icon:
            h.pack_start(Icon(icon, 16), False, False, 0)
        if markup:
            lb = Gtk.Label()
            lb.set_markup(markup)
            h.pack_start(lb, False, False, 0)
        if text:
            h.pack_start(Gtk.Label(label=text), False, False, 0)
        b.add(h)
        b.set_tooltip_text(tip)
        if cb:
            b.connect("clicked", lambda w: (cb(), self._refocus()))
        box.pack_start(b, False, False, 0)
        return b

    def _refocus(self):
        if self.editing:
            self.editor.grab_focus()
        else:
            self.da.grab_focus()

    def menu_btn(self, box, icon, tip, items_fn, text=None):
        b = self.rbtn(box, icon, tip, None, text=text)
        b.connect("clicked", lambda w: self._popup(items_fn(), w))
        return b

    def _popup(self, items, w):
        m = build_menu(items)
        m.connect("deactivate", lambda *_: GLib.idle_add(lambda: (self._refocus(), False)[1]))
        popup_menu(m, w)

    def _layout_items(self, cb):
        return [{"label": n, "icon": "layout", "cb": lambda k=k: cb(k)} for k, n in LAYOUTS]

    def rib_home(self, box):
        self.rbtn(box, "paste", "Paste (Ctrl+V)", self.paste)
        self.rbtn(box, "cut", "Cut (Ctrl+X)", lambda: self.copy_shapes(cut=True))
        self.rbtn(box, "copy", "Copy (Ctrl+C)", self.copy_shapes)
        self.rsep(box)
        self.menu_btn(box, "newslide", "New slide (Ctrl+M)", lambda: self._layout_items(self.new_slide), text="New Slide")
        self.menu_btn(box, "layout", "Slide layout", lambda: self._layout_items(self.set_layout))
        self.rbtn(box, "dup", "Duplicate slide", self.duplicate_slide)
        self.rsep(box)
        self.font_cb = Gtk.ComboBoxText.new_with_entry()
        add_cls(self.font_cb, "fontbox")
        for f in FONTS:
            self.font_cb.append_text(f)
        self.font_cb.get_child().set_width_chars(12)
        self.font_cb.connect("changed", self._font_changed)
        box.pack_start(self.font_cb, False, False, 0)
        self.size_cb = Gtk.ComboBoxText.new_with_entry()
        add_cls(self.size_cb, "sizebox")
        for s in SIZES:
            self.size_cb.append_text(str(s))
        self.size_cb.get_child().set_width_chars(3)
        self.size_cb.connect("changed", self._size_changed)
        box.pack_start(self.size_cb, False, False, 4)
        self.rbtn(box, None, "Increase font size", lambda: self.grow_font(1), markup="A<sup>+</sup>")
        self.rbtn(box, None, "Decrease font size", lambda: self.grow_font(-1), markup="A<sup>-</sup>")
        self.fmt_btns = {}
        for key, mk, tip in (("b", "<b>B</b>", "Bold (Ctrl+B)"), ("i", "<i>I</i>", "Italic (Ctrl+I)"),
                             ("u", "<u>U</u>", "Underline (Ctrl+U)")):
            self.fmt_btns[key] = self.rbtn(box, None, tip, lambda k=key: self.toggle(k), markup=mk)
        self.menu_btn(box, "fcolor", "Font color", lambda: [{"label": "Automatic", "cb": lambda: self.set_prop("color", None)}, None] +
                      [{"label": c, "pixbuf": self._swatch(c), "cb": lambda c=c: self.set_prop("color", c)} for c in TEXT_COLORS])
        self.rsep(box)
        self.bul_btn = self.rbtn(box, "bullets", "Bullets", lambda: self.toggle("bullets"))
        self.align_btns = {}
        for a in ("left", "center", "right"):
            self.align_btns[a] = self.rbtn(box, "al-" + a, f"Align {a}", lambda a=a: self.set_prop("align", a))
        self.rsep(box)
        self.menu_btn(box, "fill", "Shape fill", self._fill_items)
        self.menu_btn(box, "outline", "Shape outline", self._line_items)
        self.rbtn(box, "tofront", "Bring to front", lambda: self.arrange("front"))
        self.rbtn(box, "toback", "Send to back", lambda: self.arrange("back"))

    def _fill_items(self):
        th = self.theme
        cols = [th["accent"], th["accent2"], "#FFFFFF", "#000000", "#7F7F7F", "#C00000", "#FFC000", "#00B050", "#0070C0",
                "#7030A0", "#F4B183", "#DEEBF7"]
        return [{"label": "No fill", "cb": lambda: self.set_prop("fill", None)}, None] + \
            [{"label": c, "pixbuf": self._swatch(c), "cb": lambda c=c: self.set_prop("fill", c)} for c in cols]

    def _line_items(self):
        th = self.theme
        cols = [darker(th["accent"]), th["accent"], "#000000", "#FFFFFF", "#7F7F7F", "#C00000", "#00B050", "#0070C0"]
        return [{"label": "No outline", "cb": lambda: self.set_prop("line", None)}, None] + \
            [{"label": c, "pixbuf": self._swatch(c), "cb": lambda c=c: self.set_prop("line", c)} for c in cols] + \
            [None, {"label": "Weight", "sub": [{"label": f"{w:g} pt", "cb": lambda w=w: self.set_prop("line_w", w)}
                                               for w in (0.75, 1, 1.5, 2.25, 3, 4.5, 6)]}]

    def rib_insert(self, box):
        self.menu_btn(box, "newslide", "New slide", lambda: self._layout_items(self.new_slide), text="New Slide")
        self.rsep(box)
        self.rbtn(box, "image", "Insert a picture", self.insert_picture, text="Pictures")
        self.menu_btn(box, "shapes", "Insert a shape", lambda: [{"label": n, "cb": lambda k=k: self._start_draw(k)} for k, n in SHAPES],
                      text="Shapes")
        self.rbtn(box, "textbox", "Insert a text box", self.insert_textbox, text="Text Box")
        self.rbtn(box, "table", "Insert a table", self.insert_table, text="Table")
        self.rsep(box)
        self.rbtn(box, "history", "Insert the date", self._insert_date, text="Date")

    def _insert_date(self):
        now = datetime.datetime.now()
        text = now.strftime("%B %d, %Y").replace(" 0", " ")
        if self.editing:
            self.editor.get_buffer().insert_at_cursor(text)
            return
        self.push_undo()
        W, H = self.pres["w"], self.pres["h"]
        s = text_shape(None, W - 300, H - 50, 260, 36, 14, "right")
        s["text"] = text
        self.slide["shapes"].append(s)
        self.sel = [s]
        self.changed()

    def rib_design(self, box):
        self.theme_btns = {}
        for key, th in THEMES.items():
            b = Gtk.Button()
            add_cls(b, "themeb")
            b.set_can_focus(False)
            da = Gtk.DrawingArea()
            da.set_size_request(64, 36)
            da.connect("draw", self._draw_theme_chip, th)
            b.add(da)
            b.set_tooltip_text(th["name"])
            b.connect("clicked", lambda w, k=key: self.set_theme(k))
            box.pack_start(b, False, False, 1)
            self.theme_btns[key] = b
        self.rsep(box)
        self.menu_btn(box, "bg", "Background color", lambda: [
            {"label": "Reset to theme", "cb": lambda: self.set_bg(None)},
            {"label": "Apply this slide's background to all", "cb": lambda: self.set_bg(self.slide.get("bg"), True)}, None] +
            [{"label": c, "pixbuf": self._swatch(c), "cb": lambda c=c: self.set_bg(c)}
             for c in ("#FFFFFF", "#F2F2F2", "#000000", "#1E1E1E", "#1F3864", "#0F5E63", "#385723", "#7F6000", "#843C0C", "#3B1F3F")],
            text="Background")
        self.menu_btn(box, "layout", "Slide size", lambda: [
            {"label": "Widescreen (16:9)", "checked": self.pres["w"] > 900, "cb": lambda: self.set_size(True)},
            {"label": "Standard (4:3)", "checked": self.pres["w"] < 900, "cb": lambda: self.set_size(False)}], text="Slide Size")

    def _draw_theme_chip(self, da, cr, th):
        a = da.get_allocation()
        cr.set_source_rgb(*hex_rgb(th["bg"]))
        cr.rectangle(0, 0, a.width, a.height)
        cr.fill()
        cr.set_source_rgb(*hex_rgb(th["accent"]))
        cr.rectangle(0, a.height - 6, a.width / 2, 6)
        cr.fill()
        cr.set_source_rgb(*hex_rgb(th["accent2"]))
        cr.rectangle(a.width / 2, a.height - 6, a.width / 2, 6)
        cr.fill()
        lay = PangoCairo.create_layout(cr)
        d = Pango.FontDescription(font_for(th["tfont"]))
        d.set_absolute_size(15 * Pango.SCALE)
        lay.set_font_description(d)
        lay.set_text("Aa", -1)
        tw, tht = lay.get_pixel_size()
        cr.set_source_rgb(*hex_rgb(th["title"]))
        cr.move_to((a.width - tw) / 2, (a.height - 6 - tht) / 2)
        PangoCairo.show_layout(cr, lay)
        cr.set_source_rgba(1, 1, 1, 0.25)
        cr.rectangle(0.5, 0.5, a.width - 1, a.height - 1)
        cr.set_line_width(1)
        cr.stroke()

    def rib_show(self, box):
        self.rbtn(box, "show", "Start from the first slide (F5)", lambda: self.slideshow(0), text="From Beginning")
        self.rbtn(box, "play", "Start from this slide (Shift+F5)", lambda: self.slideshow(self.cur), text="From Current Slide")
        self.rsep(box)
        self.fade_btn = self.rbtn(box, None, "Fade between slides during the show", self._toggle_fade, text="Fade transition")

    def _toggle_fade(self):
        self.pres["fade"] = not self.pres.get("fade", True)
        self.mark_dirty()
        self._update_ribbon_state()

    def rib_view(self, box):
        self.notes_btn = self.rbtn(box, "notes", "Show the notes pane", self.toggle_notes, text="Notes")
        self.rsep(box)
        self.rbtn(box, "pdf", "Export the slides as a PDF", self.export_pdf, text="Export PDF")

    def toggle_notes(self):
        v = not self.notes_box.get_visible()
        self.notes_box.set_visible(v)
        self.st["notes"] = v
        self._update_ribbon_state()

    @staticmethod
    def _swatch(c):
        pb = GdkPixbuf.Pixbuf.new(GdkPixbuf.Colorspace.RGB, False, 8, 16, 16)
        r, g, b = (int(c[i:i + 2], 16) for i in (1, 3, 5))
        pb.fill((r << 24) | (g << 16) | (b << 8) | 0xFF)
        return pb

    def _font_changed(self, cb):
        if self.updating or cb.get_active() < 0:
            return
        self.set_prop("font", cb.get_active_text())

    def _size_changed(self, cb):
        if self.updating:
            return
        try:
            v = float((cb.get_active_text() or "").strip())
        except ValueError:
            return
        if cb.get_active() < 0 and not cb.get_child().has_focus():
            return
        if 4 <= v <= 400:
            self.set_prop("size", v)

    def _update_ribbon_state(self):
        t = self.targets()
        s = t[0] if t else None
        on = lambda b, v: (b.get_style_context().add_class if v else b.get_style_context().remove_class)("on")  # noqa: E731
        for k, b in self.fmt_btns.items():
            on(b, s is not None and s.get(k))
        for a, b in self.align_btns.items():
            on(b, s is not None and s.get("align", "left") == a)
        on(self.bul_btn, s is not None and s.get("bullets"))
        on(self.fade_btn, self.pres.get("fade", True))
        on(self.notes_btn, self.notes_box.get_visible())
        for k, b in self.theme_btns.items():
            on(b, k == self.pres["theme"])
        self.updating = True
        th = self.theme
        if s is not None:
            self.font_cb.get_child().set_text(s.get("font") or (th["tfont"] if s.get("role") == "title" else th["bfont"]))
            self.size_cb.get_child().set_text(f"{s.get('size', 18):g}")
        else:
            self.font_cb.get_child().set_text("")
            self.size_cb.get_child().set_text("")
        self.updating = False

    # ================================================================ notes
    def _notes_changed(self, buf):
        if self.updating:
            return
        self.slide["notes"] = buf.get_text(buf.get_start_iter(), buf.get_end_iter(), False)
        self.notes_hint.set_visible(not self.slide["notes"])
        self.mark_dirty()

    # ================================================================ context menu
    def shape_menu(self, ev):
        has = bool(self.sel)
        items = [{"label": "Cut", "icon": "cut", "accel": "Ctrl+X", "cb": lambda: self.copy_shapes(cut=True), "sensitive": has},
                 {"label": "Copy", "icon": "copy", "accel": "Ctrl+C", "cb": self.copy_shapes, "sensitive": has},
                 {"label": "Paste", "icon": "paste", "accel": "Ctrl+V", "cb": self.paste},
                 {"label": "Duplicate", "icon": "dup", "accel": "Ctrl+D", "cb": self.duplicate_shapes, "sensitive": has}, None]
        if has:
            items += [{"label": "Bring to front", "icon": "tofront", "cb": lambda: self.arrange("front")},
                      {"label": "Send to back", "icon": "toback", "cb": lambda: self.arrange("back")},
                      {"label": "Fill", "icon": "fill", "sub": self._fill_items()},
                      {"label": "Outline", "icon": "outline", "sub": self._line_items()}, None,
                      {"label": "Delete", "icon": "delete", "cb": lambda: self._canvas_key(None, _FakeKey(Gdk.KEY_Delete))}]
        else:
            items += [{"label": "Layout", "sub": self._layout_items(self.set_layout)},
                      {"label": "New slide", "icon": "newslide", "cb": self.new_slide}]
        popup_menu(build_menu(items), self.da, ev)

    # ================================================================ slide show
    def slideshow(self, start):
        if self.editing:
            self.commit_edit()
        ShowWindow(self, start)

    # ================================================================ files
    def file_menu(self, w):
        recent = [p for p in self.st.get("recent", []) if os.path.exists(p)][:8]
        items = [{"label": "New", "icon": "new", "accel": "Ctrl+N", "cb": lambda: self.app.new_window()},
                 {"label": "Open...", "icon": "open", "accel": "Ctrl+O", "cb": self.open_dialog},
                 {"label": "Save", "icon": "save", "accel": "Ctrl+S", "cb": self.save},
                 {"label": "Save as...", "icon": "save", "accel": "Ctrl+Shift+S", "cb": self.save_as},
                 {"label": "Export as PDF...", "icon": "pdf", "accel": "Ctrl+P", "cb": self.export_pdf}, None]
        if recent:
            items.append({"label": "Recent", "sub": [{"label": os.path.basename(p), "cb": lambda p=p: self.open_path(p)} for p in recent]})
            items.append(None)
        items.append({"label": "Close", "icon": "close", "accel": "Ctrl+W", "cb": self.close})
        popup_menu(build_menu(items), w)

    def load_pres(self, pres, path):
        self.pres = pres
        self.path = path
        self.sel = []
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.dirty = False
        self.build_thumbs()
        self.goto(0)
        self.update_title()
        self._update_ribbon_state()

    def is_blank(self):
        return self.path is None and not self.dirty

    def update_title(self):
        name = os.path.basename(self.path) if self.path else "Presentation1"
        mod = "*" if self.dirty else ""
        self.set_title(f"{mod}{name} - PowerPoint")
        self.title_lb.set_text(f"{mod}{name} - PowerPoint")

    def open_dialog(self):
        dlg = Gtk.FileChooserNative.new("Open", self, Gtk.FileChooserAction.OPEN, "Open", "Cancel")
        for name, pats in (("PowerPoint presentations (.pptx)", ["*.pptx"]), ("All files", ["*"])):
            f = Gtk.FileFilter()
            f.set_name(name)
            for p in pats:
                f.add_pattern(p)
            dlg.add_filter(f)
        dlg.set_current_folder(os.path.dirname(self.path) if self.path else os.path.join(HOME, "Documents"))
        if dlg.run() == Gtk.ResponseType.ACCEPT:
            path = dlg.get_filename()
            if self.is_blank():
                self.open_path(path)
            else:
                self.app.new_window(path)

    def open_path(self, path):
        try:
            pres = load_pptx(path)
        except Exception as e:
            info_dialog(self, "Couldn't open this file", f"{os.path.basename(path)}\n{e}")
            return
        self.load_pres(pres, path)
        self._remember(path)

    def _remember(self, path):
        r = [p for p in self.st.get("recent", []) if p != path]
        self.st["recent"] = [path] + r[:14]
        self.st.save()
        try:
            Gtk.RecentManager.get_default().add_item(GLib.filename_to_uri(path, None))
        except GLib.Error:
            pass

    def save(self):
        if self.editing:
            self.commit_edit()
        if not self.path or not self.path.lower().endswith(".pptx"):
            return self.save_as()
        return self.write(self.path)

    def save_as(self):
        if self.editing:
            self.commit_edit()
        dlg = Gtk.FileChooserNative.new("Save as", self, Gtk.FileChooserAction.SAVE, "Save", "Cancel")
        dlg.set_do_overwrite_confirmation(True)
        f = Gtk.FileFilter()
        f.set_name("PowerPoint presentation (.pptx)")
        f.add_pattern("*.pptx")
        dlg.add_filter(f)
        dlg.set_current_folder(os.path.dirname(self.path) if self.path else os.path.join(HOME, "Documents"))
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "Presentation1"
        dlg.set_current_name(base + ".pptx")
        if dlg.run() != Gtk.ResponseType.ACCEPT:
            return False
        path = dlg.get_filename()
        if not path.lower().endswith(".pptx"):
            path += ".pptx"
        return self.write(path)

    def write(self, path):
        try:
            tmp = path + ".saving"
            save_pptx(self.pres, tmp)
            os.replace(tmp, path)
        except Exception as e:
            info_dialog(self, "Couldn't save", str(e))
            return False
        self.path = path
        self.dirty = False
        self._remember(path)
        self.update_title()
        return True

    def export_pdf(self):
        if self.editing:
            self.commit_edit()
        dlg = Gtk.FileChooserNative.new("Export as PDF", self, Gtk.FileChooserAction.SAVE, "Export", "Cancel")
        dlg.set_do_overwrite_confirmation(True)
        dlg.set_current_folder(os.path.dirname(self.path) if self.path else os.path.join(HOME, "Documents"))
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "Presentation1"
        dlg.set_current_name(base + ".pdf")
        if dlg.run() != Gtk.ResponseType.ACCEPT:
            return
        path = dlg.get_filename()
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            render_pdf(self.pres, path)
        except Exception as e:
            info_dialog(self, "Couldn't export", str(e))
            return
        confirm(self, "Exported", os.path.basename(path) + " was saved.", "Open it",
                lambda ok: ok and subprocess.Popen(["xdg-open", path], start_new_session=True))

    def _cfg(self, *a):
        if not self.is_maximized():
            self.st["width"], self.st["height"] = self.get_size()
        self.st["max"] = self.is_maximized()
        if self.editing:
            GLib.idle_add(lambda: (self._style_editor(), False)[1])
        return False

    def _on_close(self, *a):
        self.st.save()
        if self.editing:
            self.commit_edit()
        if not self.dirty:
            return False
        name = os.path.basename(self.path) if self.path else "Presentation1"
        dlg = WinDialog(self, "PowerPoint", f"Do you want to save your changes to {name}?", primary="Save", secondary="Don't save")
        cancel = Gtk.Button(label="Cancel")
        add_cls(cancel, "secondary")
        cancel.connect("clicked", lambda w: (setattr(dlg, "cb", None), dlg.destroy()))
        dlg.foot.pack_start(cancel, True, True, 0)

        def done(ok):
            if ok and not self.save():
                return
            self.dirty = False
            self.destroy()
        dlg.ask(done)
        return True


class _FakeKey:
    def __init__(self, keyval):
        self.keyval = keyval
        self.state = 0


class ShowWindow(Gtk.Window):
    """Full-screen slide show."""

    def __init__(self, win, start):
        super().__init__()
        self.win = win
        self.pres = win.pres
        self.i = start
        self.done = False
        self.prev = None
        self.t0 = 0
        self.typed = ""
        self.set_transient_for(win)
        self.set_decorated(False)
        da = Gtk.DrawingArea()
        da.connect("draw", self._draw)
        self.da = da
        self.add(da)
        self.add_events(Gdk.EventMask.BUTTON_PRESS_MASK | Gdk.EventMask.POINTER_MOTION_MASK | Gdk.EventMask.SCROLL_MASK)
        self.connect("key-press-event", self._key)
        self.connect("button-press-event", self._press)
        self.connect("motion-notify-event", self._motion)
        self.connect("scroll-event", self._scroll)
        disp = win.get_display()
        mon = disp.get_monitor_at_window(win.get_window()) if win.get_window() else disp.get_primary_monitor()
        if mon is not None:
            g = mon.get_geometry()
            self.set_default_size(g.width, g.height)
            self.move(g.x, g.y)
        self.show_all()
        if mon is not None:
            self.fullscreen_on_monitor(win.get_screen(), [disp.get_monitor(i) for i in range(disp.get_n_monitors())].index(mon))
        else:
            self.fullscreen()
        self.present()
        self._hide_id = GLib.timeout_add(2500, self._hide_cursor)

    def _hide_cursor(self):
        w = self.get_window()
        if w is not None:
            w.set_cursor(Gdk.Cursor.new_for_display(self.get_display(), Gdk.CursorType.BLANK_CURSOR))
        self._hide_id = None
        return False

    def _motion(self, w, ev):
        win = self.get_window()
        if win is not None:
            win.set_cursor(None)
        if self._hide_id:
            GLib.source_remove(self._hide_id)
        self._hide_id = GLib.timeout_add(2500, self._hide_cursor)

    def _snapshot(self):
        a = self.da.get_allocation()
        surf = cairo.ImageSurface(cairo.FORMAT_RGB24, max(1, a.width), max(1, a.height))
        cr = cairo.Context(surf)
        self._paint(cr, a.width, a.height)
        return surf

    def go(self, d):
        n = len(self.pres["slides"])
        j = self.i + d
        if j < 0:
            return
        if self.pres.get("fade", True):
            self.prev = self._snapshot()
            self.t0 = time.monotonic()
            GLib.timeout_add(16, self._tick)
        if j >= n:
            self.done = True
            self.i = n
        else:
            self.done = False
            self.i = j
        self.da.queue_draw()

    def _tick(self):
        self.da.queue_draw()
        if time.monotonic() - self.t0 > 0.35:
            self.prev = None
            return False
        return True

    def _paint(self, cr, W, H):
        cr.set_source_rgb(0, 0, 0)
        cr.paint()
        if self.done:
            lay = PangoCairo.create_layout(cr)
            lay.set_font_description(Pango.FontDescription("Carlito 16"))
            lay.set_text("End of slide show, click to exit.", -1)
            cr.set_source_rgb(0.85, 0.85, 0.85)
            cr.move_to(24, 20)
            PangoCairo.show_layout(cr, lay)
            return
        pw, ph = self.pres["w"], self.pres["h"]
        sc = min(W / pw, H / ph)
        cr.save()
        cr.translate((W - pw * sc) / 2, (H - ph * sc) / 2)
        cr.scale(sc, sc)
        render_slide(cr, self.pres, self.pres["slides"][self.i])
        cr.restore()

    def _draw(self, w, cr):
        a = w.get_allocation()
        self._paint(cr, a.width, a.height)
        if self.prev is not None:
            k = min(1.0, (time.monotonic() - self.t0) / 0.35)
            cr.set_source_surface(self.prev, 0, 0)
            cr.paint_with_alpha(1 - k)

    def _key(self, w, ev):
        k = ev.keyval
        if k in (Gdk.KEY_Escape, Gdk.KEY_minus):
            self.close_show()
        elif k in (Gdk.KEY_Right, Gdk.KEY_Down, Gdk.KEY_space, Gdk.KEY_Page_Down, Gdk.KEY_n, Gdk.KEY_Return, Gdk.KEY_KP_Enter):
            if self.typed and k in (Gdk.KEY_Return, Gdk.KEY_KP_Enter):
                self.go(int(self.typed) - 1 - self.i)
                self.typed = ""
            elif self.done:
                self.close_show()
            else:
                self.go(1)
        elif k in (Gdk.KEY_Left, Gdk.KEY_Up, Gdk.KEY_Page_Up, Gdk.KEY_p, Gdk.KEY_BackSpace):
            self.go(-1)
        elif k == Gdk.KEY_Home:
            self.go(-self.i)
        elif k == Gdk.KEY_End:
            self.go(len(self.pres["slides"]) - 1 - self.i)
        elif Gdk.KEY_0 <= k <= Gdk.KEY_9:
            self.typed += chr(k)
        elif k in (Gdk.KEY_b, Gdk.KEY_period):
            pass
        return True

    def _press(self, w, ev):
        if ev.button == 1:
            if self.done:
                self.close_show()
            else:
                self.go(1)
        elif ev.button == 3:
            self.go(-1)
        return True

    def _scroll(self, w, ev):
        ok, dx, dy = ev.get_scroll_deltas()
        if not ok:
            dy = {Gdk.ScrollDirection.UP: -1, Gdk.ScrollDirection.DOWN: 1}.get(ev.direction, 0)
        if dy:
            self.go(1 if dy > 0 else -1)
        return True

    def close_show(self):
        if self._hide_id:
            GLib.source_remove(self._hide_id)
        cur = min(self.i, len(self.pres["slides"]) - 1)
        self.destroy()
        self.win.goto(cur)


class PPTApp(Gtk.Application):
    def __init__(self):
        super().__init__(application_id=APP_ID, flags=Gio.ApplicationFlags.HANDLES_COMMAND_LINE)

    def do_startup(self):
        Gtk.Application.do_startup(self)
        self.settings = JsonSettings("pi-powerpoint", DEFAULTS)
        apply_css(CSS)

    def new_window(self, path=None):
        w = PPTWindow(self, path)
        w.present()
        return w

    def do_command_line(self, cmd):
        files = [a for a in cmd.get_arguments()[1:] if not a.startswith("-")]
        if files:
            for f in files:
                p = cmd.create_file_for_arg(f).get_path()
                wins = [w for w in self.get_windows() if isinstance(w, PPTWindow) and w.is_blank()]
                if wins:
                    wins[0].open_path(p)
                    wins[0].present()
                else:
                    self.new_window(p)
        else:
            self.new_window()
        return 0


def main():
    GLib.set_prgname("pi-powerpoint")
    GLib.set_application_name("PowerPoint")
    app = PPTApp()
    signal.signal(signal.SIGINT, signal.SIG_DFL)
    return app.run(sys.argv)


if __name__ == "__main__":
    sys.exit(main())
