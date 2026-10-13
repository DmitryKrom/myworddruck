//go:generate fyne bundle -o bundled.go assets

package main

import (
	"image/color"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/theme"
)

type druckThema struct {
	fyne.Theme
}

func druckTheme() fyne.Theme {

	return &druckThema{Theme: theme.DefaultTheme()}
}

func (d *druckThema) Color(name fyne.ThemeColorName, _ fyne.ThemeVariant) color.Color {
	return d.Theme.Color(name, theme.VariantLight)
}

//	func (d *druckThema) Font(fyne.TextStyle) fyne.Resource {
//		return nil
//	}
//
//	func (d *druckThema) Icon(fyne.ThemeIconName) fyne.Resource {
//		return nil
//	}
func (d *druckThema) Size(name fyne.ThemeSizeName) float32 {
	if name == theme.SizeNameText {
		return 12
	}
	return d.Theme.Size(name)
}
