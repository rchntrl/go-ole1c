package ole1c

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type DispatchWrapper struct {
	*ole.IDispatch
}

// Подключение к базе 1С
func (wrapper DispatchWrapper) Connect(connectionString string) (conn Connection, err error) {
	connection, err := oleutil.CallMethod(wrapper.IDispatch, "Connect", connectionString)
	if err != nil {
		return
	}
	conn.VARIANT = connection
	return
}

func (wrapper DispatchWrapper) Unmarshall(template interface{}) {
	Unmarshall(wrapper.IDispatch, template)
}

func (wrapper DispatchWrapper) ToWrapperObject(template interface{}) {
	CreateOleWrapper(wrapper.IDispatch, template)
}
