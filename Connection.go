package ole1c

import (
	"github.com/go-ole/go-ole/oleutil"
	"reflect"
)

type Connection struct {
	VariantWrapper
}

func (c *Connection) CreateQuery(queryString string)  (wrapper Query, err error) {
	query, err := c.ToIDispatch().CallMethod("NewObject", "Query")
	wrapper.IDispatch = query.ToIDispatch()
	wrapper.PutProperty("text", queryString)
	return
}

func (c *Connection) NewObject(constructor string) (wrapper VariantWrapper, err error) {
	wrapper.VARIANT, err = c.ToIDispatch().CallMethod("NewObject", constructor)
	return
}

func (c *Connection) XMLString(value VariantWrapper) string {
	return c.Method("XMLString", value.Value()).Value().(string)
}

func (c *Connection) GetExternalDataProcessor(pathToFile string) VariantWrapper {
	return c.Property("ExternalDataProcessors").Method("Create", pathToFile)
}

func (c *Connection) Const(name string) (k *Constant) {
	c.Property("Константы").Property(name).Dispatched().Unmarshall(&k)
	return
}

func (c *Connection) CreateUser() VariantWrapper {
	return c.Property("InfoBaseUsers").Method("CreateUser")
}

func (c *Connection) DocumentManager(Name string) (manager DocumentManager) {
	manager.Name = Name
	manager.VariantWrapper = c.Property("Documents").Property(Name)
	manager.Dispatched().ToWrapperObject(&manager)

	var mode DocumentWriteMode
	c.Property("РежимЗаписиДокумента").Dispatched().Unmarshall(&mode)
	manager.WriteMode = &mode
	return
}

func (c *Connection) CatalogsManager(Name string) (manager CatalogsManager) {
	manager.Name = Name
	manager.VariantWrapper = c.Property("Catalogs").Property(Name)
	manager.Dispatched().ToWrapperObject(&manager)
	return
}

func (c Connection) Unmarshall(template Templatable) {
	disp := template.ToIDispatch()

	val := reflect.ValueOf(template)
	if val.Kind() != reflect.Ptr {
		panic("val must be a pointer to a struct")
	}

	val = val.Elem()

	if val.Kind() != reflect.Struct {
		panic("val must be a struct")
	}
	for i := 0; i < val.NumField(); i++ {
		field := val.Type().Field(i)
		if field.Anonymous || field.Type.Kind() != reflect.Func {
			property := val.FieldByName(field.Name)
			oleField := field.Tag.Get("db1c")
			if property.CanSet() && oleField != "" {
				// присвоение значения поля
				v := oleutil.MustGetProperty(disp, oleField)
				oleFunc := field.Tag.Get("func1c")
				if oleFunc != "" {
					// выполнить 1C функцию из Connection и присвоить его значение
					v = oleutil.MustCallMethod(c.ToIDispatch(), oleFunc, v.Value())
				}
				switch field.Type.Name()  {
				case "string":
					property.SetString(v.ToString())
				case "bool":
					property.SetBool(v.Value().(bool))
				default:
					property.Set(reflect.ValueOf(v.Value()))
					continue
				}
			}
		}
	}
}
