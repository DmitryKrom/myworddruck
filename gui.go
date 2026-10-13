package main

import (
	"image/color"
	"log"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
)

type gui struct {
	win fyne.Window
}

func makeBanner() fyne.CanvasObject {
	toolbar := widget.NewToolbar(
		widget.NewToolbarAction(theme.HomeIcon(), func() {}),
		widget.NewToolbarAction(theme.CancelIcon(), func() {}),
		widget.NewToolbarAction(theme.SearchIcon(), func() {}),
	)
	logo := canvas.NewImageFromResource(nil)
	logo.FillMode = canvas.ImageFillContain
	return container.NewStack(toolbar, container.NewPadded(logo))
}
func (g *gui) makeGUI() fyne.CanvasObject {
	top := makeBanner()
	leftSide := widget.NewLabel("left")
	middleContent := canvas.NewRectangle(color.Gray{Y: 0xee})
	rightSide := widget.NewLabel("right")

	deviders := [3]fyne.CanvasObject{
		widget.NewSeparator(),
		widget.NewSeparator(),
		widget.NewSeparator(),
	}
	objs := []fyne.CanvasObject{middleContent, top, leftSide, rightSide, deviders[0], deviders[1], deviders[2]}
	return container.New(newDruckLayout(top, leftSide, rightSide, middleContent, deviders), objs...)
}

func (g *gui) openProject() {
	dialog.ShowFolderOpen(
		func(dir fyne.ListableURI, err error) {
			log.Println("Dir opened: ", dir)
		}, g.win)
}
