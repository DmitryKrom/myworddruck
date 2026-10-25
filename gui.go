package main

import (
	"fmt"
	"image/color"
	"log"
	"os"
	"path/filepath"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/data/binding"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
)

type gui struct {
	win      fyne.Window
	folder   *widget.Label
	title    binding.String
	files    *widget.List
	fileList binding.StringList
}

func (g *gui) makeBanner() fyne.CanvasObject {
	toolbar := widget.NewToolbar(
		widget.NewToolbarAction(theme.HomeIcon(), func() { g.openProjectDialog() }),
		widget.NewToolbarAction(theme.CancelIcon(), func() { g.win.Close() }),
	)
	logo := canvas.NewImageFromResource(nil)
	logo.FillMode = canvas.ImageFillContain
	return container.NewStack(toolbar, container.NewPadded(logo))
}
func (g *gui) makeGUI() fyne.CanvasObject {
	top := g.makeBanner()
	folderWahl := widget.NewLabel("")
	g.folder = folderWahl
	druckerLable := widget.NewLabel("Ziel")
	druckerCombobox := makeCombobox()
	druckSeiten := widget.NewLabel("Seiten")
	seitenEntry := widget.NewEntry()
	seitenEntry.PlaceHolder = "Nur Seiten 1 und 2 werden gedruckt. Änderung nicht möglich!!!"
	seitenEntry.Disable()
	g.fileList = binding.NewStringList()
	g.files = widget.NewList(
		func() int { return g.fileList.Length() },
		func() fyne.CanvasObject { return widget.NewLabel("") },
		func(lii widget.ListItemID, co fyne.CanvasObject) {
			val, _ := g.fileList.Get()
			co.(*widget.Label).SetText(val[lii])
		},
	)
	middleContent := container.NewStack(canvas.NewRectangle(color.Gray{Y: 0xee}), g.files)
	druckButton := widget.NewButton("Drucken", func() {
		if g.files.Length() == 0 {
			dialog.ShowError(fmt.Errorf("Es sind keine Dateien in dem Ordner !!!"), g.win)
			return
		}
		drucker := druckerCombobox.Text
		seiten := seitenEntry.Text
		// var von, bis string

		if strings.Contains(seiten, "-") {
			// von, bis, _ = strings.Cut(seiten, "-")
		}
		g.print(drucker, "1", "2")
	})
	druckCont := container.NewGridWithRows(20, druckerLable, druckerCombobox, druckSeiten, seitenEntry, druckButton)
	// leftSide := widget.NewLabel("left")
	// rightSide := widget.NewLabel("right")
	deviders := [3]fyne.CanvasObject{
		widget.NewSeparator(),
		widget.NewSeparator(),
		widget.NewSeparator(),
	}
	objs := []fyne.CanvasObject{top, folderWahl, druckCont, middleContent, deviders[0], deviders[1], deviders[2]}
	return container.New(newDruckLayout(top, folderWahl, druckCont, middleContent, deviders), objs...)
}

func makeCombobox() *widget.SelectEntry {
	druckerListe, err := listPrinters()
	var printerListe []string
	if err != nil {
		fmt.Println("[Error] Beim Druckeraulisten: ", err)
	}
	for i := range druckerListe {
		printerListe = append(printerListe, druckerListe[i].name)
	}
	sel := widget.NewSelectEntry(printerListe)
	return sel
}

func (g *gui) openProjectDialog() {
	fmt.Println("Dialog !!!")
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
		g.folder.SetText(name)
		filesURIs, err := dir.List()
		if err != nil {
			fmt.Println("[Error] filesURIs : ", err)
		}
		g.fileList.Set([]string{})
		for i := range filesURIs {
			di, _ := os.Stat(filesURIs[i].Path())
			fmt.Println(filesURIs[i].Path(), " ", filesURIs[i].Extension() != ".docx")
			if !di.IsDir() {
				if filesURIs[i].Extension() == ".docx" || filesURIs[i].Extension() == ".doc" {
					g.fileList.Append(filesURIs[i].Path())
				}
			}
		}
		g.files.Refresh()
	} else {
		dialog.ShowError(fmt.Errorf("Das ist eine Datei! Wählen Sie ein Verzeichnis!"), g.win)
		return
	}

}
func filePathWalkDir(root string) ([]string, error) {
	entries, err := os.ReadDir(root)
	if err != nil {
		log.Fatal(err)
	}

	var files []string
	for _, v := range entries {
		if v.IsDir() {
			continue
		}
		if filepath.Ext(v.Name()) != ".md" {
			continue
		}
		files = append(files, v.Name())
	}

	return files, nil
}
