package main

import (
	"fmt"
	"log"
	"os"
	"os/exec"
	"runtime"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
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

func main() {
	fmt.Println("Starte mein Druck !!!")
	// Drucker und Verzeichnis auswählen
	druckerListe, err := listPrinters()
	if len(druckerListe) == 0 {
		fmt.Println("Keine Drucker gefunden !!!")
	}
	_, _ = verzeichnisAndDruckerWahl(druckerListe)

	os.Exit(1)

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
	word, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal(err)
	}
	defer word.Release()

	_, _ = oleutil.PutProperty(word, "Visible", false)

	// =========================================================================
	// NEU: DRUCKERAUSWAHL FESTLEGEN
	// Ersetze diesen Namen exakt durch den Namen deines Druckers aus Windows
	// =========================================================================
	wunschDrucker := ""

	fmt.Printf("Wechsle aktiven Drucker zu: %s...\n", wunschDrucker)
	_, err = oleutil.PutProperty(word, "ActivePrinter", wunschDrucker)
	if err != nil {
		log.Printf("Warnung: Konnte Drucker nicht explizit setzen (Prüfe den Namen): %v", err)
		// Falls der Name falsch ist, nutzt Word standardmäßig den Windows-Standarddrucker
	}
	fmt.Printf("Wechsle aktiven Drucker 2 zu: %s...\n", wunschDrucker)
	// 3. Dokument öffnen
	documents := oleutil.MustGetProperty(word, "Documents").ToIDispatch()
	defer documents.Release()

	docPath := ``
	docVariant, err := oleutil.CallMethod(documents, "Open", docPath)
	if err != nil {
		log.Fatalf("Konnte Dokument nicht öffnen: %v", err)
	}
	fmt.Printf("Wechsle aktiven Drucker 3 zu: %s...\n", wunschDrucker)
	doc := docVariant.ToIDispatch()
	defer doc.Release()

	// 4. Nur die ersten zwei Seiten drucken
	wdPrintFromTo := 3

	fmt.Println("Sende Seiten 1-2 an den ausgewählten Drucker...")
	_, err = oleutil.CallMethod(doc, "PrintOut",
		false,         // Background: false wartet, bis die Daten im Spooler sind
		false,         // Append
		wdPrintFromTo, // Range: "Von-Bis" Bereich nutzen
		"",            // OutputFileName
		"1",           // From: Startseite
		"2",           // To: Endseite
	)

	if err != nil {
		log.Printf("Fehler beim Drucken: %v", err)
	} else {
		fmt.Println("Druckauftrag erfolgreich übergeben.")
	}

	time.Sleep(2 * time.Second)

	// 5. Aufräumen
	_, _ = oleutil.CallMethod(doc, "Close", 0)
	_, _ = oleutil.CallMethod(word, "Quit")
	fmt.Println("Word erfolgreich geschlossen.")
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

func verzeichnisAndDruckerWahl(d []printer) (string, string) {
	var verz verzeichnis

	app := app.NewWithID("druck")
	window := app.NewWindow("Ausdruck")

	window.Resize(fyne.Size{Width: 500, Height: 600})
	// window.SetContent(widget.NewLabel("Wählen Sie einen Drucker !"))
	druckerName, err := fillUpDruckerButtons(window, d)
	if err != nil {
		return err.Error(), " "
	}
	window.ShowAndRun()

	return druckerName, verz.pfad
}

func fillUpDruckerButtons(w fyne.Window, druckerListe []printer) (string, error) {
	if len(druckerListe) < 1 {
		return "", fmt.Errorf("Druckerliste ist leer !!!")
	}
	var druckerName string
	var objcts = make([]fyne.CanvasObject, 1)
	objcts[0] = widget.NewLabel("Wählen Sie einen Drucker !")
	for i := range druckerListe {
		btn := widget.NewButton(druckerListe[i].name, func() {
			druckerName = druckerListe[i].name
			w.Close()
		})
		objcts = append(objcts, btn)
	}

	w.SetContent(container.NewVBox(objcts...))
	return druckerName, nil
}
