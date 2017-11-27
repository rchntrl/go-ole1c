package ole1c

import (
	"github.com/go-ole/go-ole/oleutil"
)

type ValueTable struct {
	DispatchWrapper
}

func (table *ValueTable) Count() int {
	v := oleutil.MustCallMethod(table.IDispatch, "count")
	return int(v.Val)
}

func (table *ValueTable) Get(index int) (record VariantWrapper) {
	record.VARIANT = oleutil.MustCallMethod(table.IDispatch, "get", index)
	return
}
