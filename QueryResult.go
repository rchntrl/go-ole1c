package ole1c

import "github.com/go-ole/go-ole/oleutil"

type QueryResult struct {
	DispatchWrapper
}

// Метод Выбрать() ВыборкаРезультатаЗапроса
func (wrapper QueryResult) Choose() (selection QueryResultSelection) {
	selection.IDispatch = oleutil.MustCallMethod(wrapper.IDispatch,"Choose").ToIDispatch()
	return
}

// Метод Выгрузить() ТаблицаЗначений
func (wrapper QueryResult) Unload() (table ValueTable) {
	table.IDispatch = oleutil.MustCallMethod(wrapper.IDispatch, "Unload").ToIDispatch()
	return
}
