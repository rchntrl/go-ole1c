package ole1c

import (
	"time"
)

type DocumentWriteMode struct {
	Write interface{} `db1c:"Write"`
	UndoPosting interface{} `db1c:"UndoPosting"`
	Posting interface{} `db1c:"Posting"`
}

type DocumentManager struct {
	VariantWrapper
	Name string
	WriteMode *DocumentWriteMode
}

func (manager DocumentManager) Create() (de DocumentObject) {
	wrapper := manager.Method("CreateDocument").Dispatched()
	wrapper.Unmarshall(&de)
	return
}

func (manager DocumentManager) FindByNumber(number interface{}) (ref DocumentRef) {
	ref.VariantWrapper = manager.Method("FindByNumber", number)
	wrapperDoc := ref.Dispatched()
	wrapperDoc.ToWrapperObject(&ref.DocumentReadable)
	wrapperDoc.Unmarshall(&ref.DocumentReadable)
	return
}

func (manager DocumentManager) MarkAsDeleted(doc *DocumentObject) {
	doc.SetDeletionMark(true)
	doc.Write(manager.WriteMode.Write)
}

func (manager DocumentManager) UnDelete(object *DocumentObject) {
	object.SetDeletionMark(false)
	object.Write(manager.WriteMode.Write)
}

func (manager DocumentManager) Post(object DocumentObject) {
	object.Write(manager.WriteMode.Posting)
}

func (manager DocumentManager) Save(object DocumentObject) {
	object.Write(manager.WriteMode.Write)
}

func (manager DocumentManager) UnPost(object DocumentObject) {
	object.Write(manager.WriteMode.UndoPosting)
}

type DocumentReadable struct {
	Number interface{} `db1c:"Номер"`
	Date time.Time `db1c:"Дата"`
	DeletionMark bool `db1c:"ПометкаУдаления"`
	Posted bool `db1c:"Проведен"`
}

type DocumentRef struct {
	VariantWrapper
	DocumentReadable
}

func (ref DocumentRef) GetWritable() (object DocumentObject) {
	ref.Method("ПолучитьОбъект").Dispatched().ToWrapperObject(&object)
	return
}

type DocumentObject struct {
	// Ссылка
	Ref func() interface{} `ole:"get"`
	Number func() string `ole:"get"`
	Date func() time.Time `ole:"get"`
	// Проведен
	Posted func() bool `ole:"get"`
	DeletionMark func() bool `ole:"get"`
	// Пометка удаления
	SetNumber func(number interface{}) `ole:"set" ole.name:"Number"`
	SetDeletionMark func(deletionMark bool) `ole:"set" ole.name:"ПометкаУдаления"`
	SetNewNumber func() `ole:"method"`
	Write func(mode interface{}) `ole:"method"`
}
