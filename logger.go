package main

import (
	"io"
	"log"
	"os"
)

var Logger *log.Logger

func InitLogger() error {

	file, err := os.OpenFile(
		"druck.log",
		os.O_CREATE|os.O_APPEND|os.O_WRONLY,
		0666,
	)

	if err != nil {
		return err
	}

	Logger = log.New(
		io.MultiWriter(os.Stdout, file),
		"[Druck App] ",
		log.Ldate|log.Ltime|log.Lshortfile,
	)

	return nil
}
