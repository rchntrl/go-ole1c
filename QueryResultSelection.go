package ole1c

import (
	"github.com/go-ole/go-ole/oleutil"
)

type QueryResultSelection struct {
	DispatchWrapper
}

func (selection *QueryResultSelection) GetField(fieldName string) (v VariantWrapper) {
	v.VARIANT = oleutil.MustGetProperty(selection.IDispatch, fieldName)
	return
}

func (selection *QueryResultSelection) Count() int {
	v := oleutil.MustCallMethod(selection.IDispatch, "count")
	return int(v.Val)
}

func (selection *QueryResultSelection) Next() bool  {
	v := oleutil.MustCallMethod(selection.IDispatch, "next")
	return v.Val != 0
}

func (selection *QueryResultSelection) Previous() bool  {
	v := oleutil.MustCallMethod(selection.IDispatch, "previous")
	return v.Val != 0
}
