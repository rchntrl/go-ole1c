package ole1c

import (
	"github.com/go-ole/go-ole/oleutil"
)

type Query struct {
	DispatchWrapper
}

func (query *Query) SetParameter(name string, value interface{}) {
	oleutil.CallMethod(query.IDispatch,"setParameter", name, value)
}

func (query *Query) Execute() (result QueryResult) {
	 r := oleutil.MustCallMethod(query.IDispatch,"Execute")
	result.IDispatch = r.ToIDispatch()
	result.ToWrapperObject(&result)
	return
}
