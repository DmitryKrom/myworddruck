package main

import (
	"fmt"
	"image/color"
	"os"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/data/binding"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
)

type gui struct {
	win   fyne.Window
	title binding.String
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
func (g *gui) makeGUI(druckerListe []printer) fyne.CanvasObject {
	top := makeBanner()
	druckerLable := widget.NewLabel("Ziel")
	druckerCombobox := makeCombobox(druckerListe)
	druckSeiten := widget.NewLabel("Seiten")
	seitenEntry := widget.NewEntry()
	seitenEntry.PlaceHolder = "z.B.: 1-2"
	directory := widget.NewLabelWithData(g.title)
	middleContent := container.NewStack(canvas.NewRectangle(color.Gray{Y: 0xee}), directory)
	druckButton := widget.NewButton("Drucken", func() {

	})
	druckCont := container.NewGridWithRows(20, druckerLable, druckerCombobox, druckSeiten, seitenEntry, druckButton)
	leftSide := widget.NewLabel("left")
	// rightSide := widget.NewLabel("right")
	deviders := [3]fyne.CanvasObject{
		widget.NewSeparator(),
		widget.NewSeparator(),
		widget.NewSeparator(),
	}
	objs := []fyne.CanvasObject{middleContent, top, leftSide, druckCont, deviders[0], deviders[1], deviders[2]}
	return container.New(newDruckLayout(top, leftSide, druckCont, middleContent, deviders), objs...)
}

func makeCombobox(d []printer) *widget.SelectEntry {
	druckerNamen := make([]string, 1)
	for i := range d {
		druckerNamen = append(druckerNamen, d[i].name)
	}
	sel := widget.NewSelectEntry(druckerNamen)
	return sel
}

func (g *gui) openProjectDialog() {
	fmt.Println("Dialog !!!")
	// dialog.ShowFileOpen(func(reader fyne.URIReadCloser, err error) {
	// 	fmt.Println(reader.URI())
	// }, g.win)
	dialog.ShowFolderOpen(
		func(dir fyne.ListableURI, err error) {
			if err != nil {
				dialog.ShowError(err, g.win)
				return
			}
			if dir == nil {
				return
			}
			g.openProject(dir)
		}, g.win)

}

func (g *gui) openProject(dir fyne.ListableURI) {
	fmt.Println("openProjekt !!!!!")
	dirInfo, err := os.Stat(dir.Path())
	if err != nil {
		dialog.ShowError(err, g.win)
		return
	}
	if dirInfo.IsDir() {
		name := dir.Name()
		g.win.SetTitle("Verzeichnis: " + name)
		g.title.Set(name)
	} else {
		dialog.ShowError(fmt.Errorf("Das ist eine Datei! Wählen Sie ein Verzeichnis!"), g.win)
		return
	}

}
