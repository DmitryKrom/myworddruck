package main

import (
	"fmt"
	"log"
	"runtime"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func main() {

	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	err := ole.CoInitialize(0)
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
