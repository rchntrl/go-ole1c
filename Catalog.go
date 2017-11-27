package ole1c

type CatalogsManager struct {
	VariantWrapper
	Name string
}

func (manager *CatalogsManager) Select(params ...interface{}) (selection QueryResultSelection) {
	selection.DispatchWrapper = manager.Method("Select", params...).Dispatched()
	return
}

func (manager *CatalogsManager) FindByCode(code interface{}) (ref CatalogRef) {
	ref.VariantWrapper = manager.Method("FindByCode", code)
	ref.Dispatched().Unmarshall(&ref)
	return
}

func (manager *CatalogsManager) FindByCodeInto(code interface{}, ref interface{}) {
	manager.Method("FindByCode", code).Dispatched().Unmarshall(&ref)
	return
}

func (manager *CatalogsManager) FindByName(name interface{}) (ref CatalogRef) {
	ref.VariantWrapper = manager.Method("FindByName", name)
	return
}

type CatalogRef struct {
	VariantWrapper
	Code interface{} `db1c:"Code"`
	Description string `db1c:"Description"`
	DeletionMark bool `db1c:"DeletionMark"`
}

func (ref *CatalogRef) GetWritable() (object CatalogObject) {
	ref.Method("ПолучитьОбъект").Dispatched().ToWrapperObject(&object)
	return
}

type CatalogObject struct {
	Ref func() interface{} `ole:"get"`
	Code func() string `ole:"get"`
	Description func() string `ole:"get"`
	DeletionMark func() bool `ole:"get"`
	SetDeletionMark func(deletionMark bool) `ole:"set" ole.name:"ПометкаУдаления"`
	Write func() `ole:"method"`
	Delete func() `ole:"method"`
}
