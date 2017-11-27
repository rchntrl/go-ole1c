package ole1c

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type VariantWrapper struct {
	*ole.VARIANT
}

// Доступ к методу COM-объекта
func (wrapper VariantWrapper) Method(name string, params ...interface{}) (v VariantWrapper) {
	v.VARIANT = oleutil.MustCallMethod(wrapper.ToIDispatch(), name, params...)
	return
}

// Доступ к свойству COM-объекта
func (wrapper VariantWrapper) Property(name string, params ...interface{}) (v VariantWrapper) {
	v.VARIANT = oleutil.MustGetProperty(wrapper.ToIDispatch(), name, params...)
	return
}

// Установить значение свойству COM-объекта
func (wrapper VariantWrapper) PutProperty(name string, params ...interface{}) {
	wrapper.VARIANT, _ = wrapper.ToIDispatch().PutProperty( name, params...)
}

// Преобразуем в обертку DispatchWrapper
func (wrapper VariantWrapper) Dispatched() DispatchWrapper {
	return DispatchWrapper{wrapper.ToIDispatch()}
}
