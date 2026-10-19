package main

import (
	"flag"
	"fmt"
	"log"
	"os/exec"
	"path/filepath"
	"runtime"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/data/binding"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type printer struct {
	name string
}
type verzeichnis struct {
	pfad string
}

type druckApp struct {
	g gui
}

func newDruckApp(name string) *druckApp {

	app := app.NewWithID(name)
	app.Settings().SetTheme(druckTheme())
	window := app.NewWindow(name)
	window.SetPadded(false)
	window.Resize(fyne.Size{Width: 1024, Height: 768})
	// // window.SetContent(widget.NewLabel("Wählen Sie einen Drucker !"))
	// druckerName, err := fillUpDruckerButtons(window, d)
	// if err != nil {
	// 	return err.Error(), " "
	// }
	ui := &gui{win: window, title: binding.NewString()}
	window.SetContent(ui.makeGUI())
	window.SetMainMenu(ui.makeMenu())
	ui.title.AddListener(binding.NewDataListener(func() {
		name, _ := ui.title.Get()
		window.SetTitle(name)
	}))
	// fmt.Println(" --> ", ui.win.Canvas().Content())
	d := &druckApp{g: *ui}
	flag.Usage = func() {
		fmt.Println("Usage: druck [project directory]")
	}
	flag.Parse()
	if len(flag.Args()) > 0 {
		dirPath := flag.Args()[0]
		absDir, err := filepath.Abs(dirPath)
		if err != nil {
			return nil
		}
		dirURI := storage.NewFileURI(absDir)
		dir, err := storage.ListerForURI(dirURI)
		if err != nil {
			fmt.Println("Error opening project", err)
			return nil
		}
		d.g.openProject(dir)
	} else {
		d.g.openProjectDialog()
	}
	return d
}
func (a *druckApp) run() {
	a.g.win.ShowAndRun()
}

func main() {
	app := newDruckApp("Druck Vertrag")
	app.run()
}

func (g *gui) print(drucker, von, bis string) {
	docListe, err := g.fileList.Get()
	if err != nil {
		fmt.Println("[Error] lesen FileListe : ", err)
	}

	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	err = ole.CoInitialize(0)
	if err != nil {
		log.Fatalf("Konnte OLE nicht initialisieren: %v", err)
	}
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("Word.Application")
	if err != nil {
		log.Fatalf("Word ist vermutlich nicht installiert: %v", err)
	}
	defer unknown.Release()

	word, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal(err)
	}
	defer word.Release()

	// Word am Ende sauber beenden (läuft VOR den Releases, da defer LIFO)
	defer func() {
		_, _ = oleutil.CallMethod(word, "Quit", 0) // 0 = wdDoNotSaveChanges
		fmt.Println("Word erfolgreich geschlossen.")
	}()

	_, _ = oleutil.PutProperty(word, "Visible", false)
	_, _ = oleutil.PutProperty(word, "DisplayAlerts", 0) // keine Dialoge

	// =========================================================================
	// NEU: DRUCKERAUSWAHL FESTLEGEN
	// Ersetze diesen Namen exakt durch den Namen deines Druckers aus Windows
	// =========================================================================

	fmt.Printf("Wechsle aktiven Drucker zu: %s...\n", drucker)
	_, err = oleutil.PutProperty(word, "ActivePrinter", drucker)
	if err != nil {
		log.Printf("Warnung: Konnte Drucker nicht explizit setzen (Prüfe den Namen): %v", err)
		// Falls der Name falsch ist, nutzt Word standardmäßig den Windows-Standarddrucker
	}
	// 3. Dokument öffnen
	documents := oleutil.MustGetProperty(word, "Documents").ToIDispatch()
	defer documents.Release()

	for _, docPath := range docListe {
		g.printOne(documents, docPath, von, bis)
	}
	fmt.Println("Word erfolgreich geschlossen.")
}

func (g *gui) printOne(documents *ole.IDispatch, docPath, von, bis string) {
	docVariant, err := oleutil.CallMethod(documents, "Open", docPath)
	if err != nil {
		log.Printf("Konnte Dokument nicht öffnen (%s): %v", docPath, err)
		return
	}
	doc := docVariant.ToIDispatch()
	defer doc.Release()

	const wdPrintFromTo = 3
	fmt.Printf("Drucke %s, Seiten %s-%s\n", docPath, von, bis)

	_, err = oleutil.CallMethod(doc, "PrintOut",
		false,         // Background: false = warten
		false,         // Append
		wdPrintFromTo, // Range
		"",            // OutputFileName
		von,           // From
		bis,           // To
	)
	if err != nil {
		log.Printf("Fehler beim Drucken: %v", err)
	} else {
		fmt.Println("Druckauftrag erfolgreich übergeben.")
	}

	// Dokument schließen ohne Speichern
	_, _ = oleutil.CallMethod(doc, "Close", 0)

}

// listPrinters zeigt betriebssystemspezifisch die verfügbaren Drucker an,
// damit der Nutzer einen gültigen Namen für -printer ermitteln kann.
func listPrinters() ([]printer, error) {
	var cmd *exec.Cmd
	switch runtime.GOOS {
	case "windows":
		cmd = exec.Command("powershell", "-NoProfile", "-Command",
			"Get-Printer | Select-Object -ExpandProperty Name")
	default: // macOS und Linux nutzen beide CUPS
		cmd = exec.Command("lpstat", "-p")
	}

	var drucker []printer
	out, err := cmd.CombinedOutput()
	if err != nil {
		return nil, fmt.Errorf("%w\n%s", err, string(out))
	}
	var drName string
	x := 0
	for i := range out {
		if string(out[i]) == "\r" {
			drName = string(out[x : i+1])
			// fmt.Println("drName: ", drName)
			var d printer
			d.name = drName
			drucker = append(drucker, d)
			i++
			x = i + 1
		}
	}
	fmt.Println("Verfügbare Drucker:")
	for i := range drucker {
		fmt.Println(drucker[i].name)
	}
	fmt.Println("#############################################")

	return drucker, nil
}

func verzeichnisAndDruckerWahl(druckerListe []printer) (string, string, error) {
	var verz verzeichnis

	app := app.NewWithID("druck")
	app.Settings().SetTheme(druckTheme())
	window := app.NewWindow("Ausdruck")
	window.SetPadded(false)
	window.Resize(fyne.Size{Width: 1024, Height: 768})
	// // window.SetContent(widget.NewLabel("Wählen Sie einen Drucker !"))
	// druckerName, err := fillUpDruckerButtons(window, d)
	// if err != nil {
	// 	return err.Error(), " "
	// }
	ui := &gui{win: window, title: binding.NewString()}
	window.SetContent(ui.makeGUI())
	window.SetMainMenu(ui.makeMenu())
	ui.title.AddListener(binding.NewDataListener(func() {
		name, _ := ui.title.Get()
		window.SetTitle(name)
	}))
	flag.Usage = func() {
		fmt.Println("Usage: druck [project directory]")
	}
	flag.Parse()
	if len(flag.Args()) > 0 {
		dirPath := flag.Args()[0]
		absDir, err := filepath.Abs(dirPath)
		if err != nil {
			return "", "", err
		}
		dirURI := storage.NewFileURI(absDir)
		dir, err := storage.ListerForURI(dirURI)
		if err != nil {
			fmt.Println("Error opening project", err)
			return "", "", err
		}
		ui.openProject(dir)
	} else {
		ui.openProjectDialog()
	}

	window.ShowAndRun()

	return verz.pfad, " ", nil
}

func (g *gui) makeMenu() *fyne.MainMenu {
	file := fyne.NewMenu("File",
		fyne.NewMenuItem("Open Project", g.openProjectDialog),
	)
	return fyne.NewMainMenu(file)
}

func fillUpDruckerButtons(w fyne.Window, druckerListe []printer) (string, error) {
	if len(druckerListe) < 1 {
		return "", fmt.Errorf("Druckerliste ist leer !!!")
	}
	var druckerName string
	var printBtns = make([]fyne.CanvasObject, 1)
	dirPath, err := filepath.Abs("C://Users")
	if err != nil {
		fmt.Println("Error resolving project path", err)
	}

	dirURI := storage.NewFileURI(dirPath)
	dir, err := storage.ListerForURI(dirURI)
	if err != nil {
		fmt.Println("Error opening project", err)
	}

	lab := widget.NewLabel(dirPath)
	fileTree := binding.NewURITree()

	files := widget.NewTreeWithData(
		fileTree,
		func(branch bool) fyne.CanvasObject {
			return widget.NewLabel("Name")
		},
		func(data binding.DataItem, branch bool, obj fyne.CanvasObject) {
			l := obj.(*widget.Label)
			u, _ := data.(binding.URI).Get()

			l.SetText(u.Name())
		},
	)

	files.OnSelected = func(id widget.TreeNodeID) {
		u, err := fileTree.GetValue(id)
		if err != nil {
			dialog.ShowError(err, nil)
			files.Unselect(id)
			return
		}
		fmt.Println(u)

		listable, err := storage.CanList(u)
		if !listable || err != nil {
			files.Unselect(id)
			return
		}

		// _, err = g.openFile(u)
		// if err != nil {
		// 	dialog.ShowError(err, g.win)
		// 	files.Unselect(id)
		// }
		fmt.Println("\t ", listable, "  u: ", u.Name())
		next, _ := storage.ListerForURI(u)
		addFilesToTree(next, fileTree, id)
		lab.SetText(filepath.Dir(u.Name()))
	}
	fileTree.Set(map[string][]string{}, map[string]fyne.URI{})
	addFilesToTree(dir, fileTree, binding.DataTreeRootID)

	printBtns[0] = widget.NewLabel("Wählen Sie einen Drucker !")
	for i := range druckerListe {
		btn := widget.NewButton(druckerListe[i].name, func() {
			druckerName = druckerListe[i].name
			w.Close()
		})
		printBtns = append(printBtns, btn)
	}
	c := container.NewVBox(
		lab,
		files,
		container.NewHBox(printBtns...),
	)
	// w.SetContent(container.NewVBox(printBtns...))
	w.SetContent(c)
	return druckerName, nil
}

func addFilesToTree(dir fyne.ListableURI, tree binding.URITree, root string) {
	items, _ := dir.List()
	for _, uri := range items {
		name := uri.Name()
		if len(name) > 0 && (name[0] == '.' || name == "go.sum") {
			continue
		}
		// pos := strings.LastIndex(name, ".gui.go")
		// if pos != -1 && pos == len(name)-7 {
		// 	continue
		// }

		nodeID := uri.String()
		tree.Append(root, nodeID, uri)

		_, err := storage.CanList(uri)
		if err != nil {
			log.Println("Failed to check for listing")
		}

	}
}
