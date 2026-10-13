package main

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/theme"
)

const sideWidth = 220
const sideHight = 150

type druckLayout struct {
	top, left, right, content fyne.CanvasObject
	deviders                  [3]fyne.CanvasObject
}

func newDruckLayout(top, left, right, content fyne.CanvasObject, deviders [3]fyne.CanvasObject) fyne.Layout {
	return &druckLayout{
		top:      top,
		left:     left,
		right:    right,
		content:  content,
		deviders: deviders,
	}
}

func (dl *druckLayout) Layout(o []fyne.CanvasObject, size fyne.Size) {
	topHight := dl.top.MinSize().Height
	dl.top.Resize(fyne.NewSize(size.Width, topHight))

	dl.left.Move(fyne.NewPos(0, topHight))
	dl.left.Resize(fyne.NewSize(sideWidth, size.Height-topHight))

	dl.right.Move(fyne.NewPos(size.Width-sideWidth, topHight))
	dl.right.Resize(fyne.NewSize(sideWidth, size.Height-topHight))

	dl.content.Move(fyne.NewPos(sideWidth, topHight))
	dl.content.Resize(fyne.NewSize(size.Width-sideWidth*2, size.Height-topHight))

	deviderThickness := theme.SeparatorThicknessSize()
	dl.deviders[0].Move(fyne.NewPos(0, topHight))
	dl.deviders[0].Resize(fyne.NewSize(size.Width, deviderThickness))

	dl.deviders[1].Move(fyne.NewPos(sideWidth, topHight))
	dl.deviders[1].Resize(fyne.NewSize(deviderThickness, size.Height-topHight))

	dl.deviders[2].Move(fyne.NewPos(size.Width-sideWidth, topHight))
	dl.deviders[2].Resize(fyne.NewSize(deviderThickness, size.Height-topHight))

}

func (dl *druckLayout) MinSize(o []fyne.CanvasObject) fyne.Size {

	borders := fyne.NewSize(
		sideHight*2,
		dl.top.MinSize().Height,
	)

	return borders.AddWidthHeight(400, 200)
}
