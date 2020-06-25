package ole1c

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"reflect"
	"runtime"
)

type Float1C interface{}

type Templatable interface {
	ToIDispatch() *ole.IDispatch
}

// CreateConnector создает COM-соединение и возвращает обертку для объекта
func CreateConnector() (conn DispatchWrapper, err error) {
	ole.CoInitialize(0)
	unknown, err := oleutil.CreateObject("V82.COMConnector")
	if err != nil {
		return
	}
	conn.IDispatch, err = unknown.QueryInterface(ole.IID_IDispatch)
	return
}

func CreateConnectorEx(p uintptr) (conn DispatchWrapper, err error) {
	ole.CoInitializeEx(p, 0)
	unknown, err := oleutil.CreateObject("V82.COMConnector")
	if err != nil {
		return
	}
	conn.IDispatch, err = unknown.QueryInterface(ole.IID_IDispatch)
	return
}

// DestroyConnector закрывает COM-соединение
func DestroyConnector() {
	ole.CoUninitialize()
}

func Unmarshall(disp *ole.IDispatch, template interface{}) {
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
				// присвоение значение поля
				typename := field.Type.Name()
				v := oleutil.MustGetProperty(disp, oleField)
				switch typename {
				case "string":
					property.SetString(v.ToString())
				case "bool":
					property.SetBool(v.Value().(bool))
				case "Float1C":
					if v.Value() == nil {
						property.Set(reflect.ValueOf(0))
					} else {
						property.Set(reflect.ValueOf(v.Value()))
					}
				case "":
					property.Set(reflect.ValueOf(v))
				default:
					property.Set(reflect.ValueOf(v.Value()))
					continue
				}
			}
		}
	}
}

func CreateOleWrapper(disp *ole.IDispatch, template interface{}) {
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
			continue
		}

		var callfunc interface{}

		switch field.Tag.Get("ole") {
		case "get":
			callfunc = oleutil.GetProperty
		case "set":
			callfunc = oleutil.PutProperty
		case "method":
			callfunc = oleutil.CallMethod
		default:
			continue
		}

		name := field.Name
		olename := field.Tag.Get("ole.name")
		if olename > "" {
			name = olename
		}
		validateInvokeFunc(val.Type().Name(), field.Name, field.Type)
		f := reflect.MakeFunc(field.Type, makeInvokeFunc(disp, callfunc, name, field.Type))
		val.FieldByName(field.Name).Set(f)
	}
}

func validateInvokeFunc(structname string, fieldname string, functype reflect.Type) {
	if functype.NumOut() > 2 {
		panic(fmt.Errorf("Error in OLE wrapper struct %s (invalid method signature <%s> for field '%s'): Method must return 0, 1, or 2 values", structname, functype, fieldname))
	}

	if functype.NumOut() == 2 {
		// function type must return (T, error) [in that order]
		valid := !isErrorType(functype.Out(0)) && isErrorType(functype.Out(1))

		if !valid {
			panic(fmt.Errorf("Error in OLE wrapper struct %s (invalid method signature <%s> for field '%s'): Method has 2 return values. They must be in the form (T, error), where T is any non-error type", structname, functype, fieldname))
		}
	}
}

func makeInvokeFunc(disp *ole.IDispatch, callfunc interface{}, name string, functype reflect.Type) func([]reflect.Value) []reflect.Value {
	return func(in []reflect.Value) []reflect.Value {

		results := make([]reflect.Value, 0)

		f := reflect.ValueOf(callfunc)

		params := make([]reflect.Value, 0)
		params = append(params, reflect.ValueOf(disp))
		params = append(params, reflect.ValueOf(name))
		params = append(params, in...)

		out := f.Call(params)
		err := out[1].Interface()

		v := out[0].Interface().(*ole.VARIANT)

		if v != nil {
			// Not sure this is a good idea, particularly for *ole.IDispatch results...
			runtime.SetFinalizer(v, func(v *ole.VARIANT) { ole.VariantClear(v) })
		}

		// If the function in the template struct returns at least one value, and the first
		// value is *not* an error type, populate the return value into that return value

		if functype.NumOut() > 0 && !isErrorType(functype.Out(0)) {
			// nil/null/empty case - return zero-value for all of these
			if v == nil || v.VT == ole.VT_NULL || v.VT == ole.VT_EMPTY || (v.VT == ole.VT_DISPATCH && v.Val == 0) {
				results = append(results, reflect.Zero(functype.Out(0)))
			} else if v.VT == ole.VT_DISPATCH {
				rdisp := v.ToIDispatch()

				rt := functype.Out(0)
				if rt.Kind() == reflect.Ptr {
					rt = rt.Elem()
				}

				wrapper := reflect.New(rt)
				CreateOleWrapper(rdisp, wrapper.Interface())

				val := wrapper.Elem()
				if functype.Out(0).Kind() == reflect.Ptr {
					val = wrapper
				}

				results = append(results, val)
			} else {
				results = append(results, reflect.ValueOf(v.Value()))
			}
		}

		// If the function in the template struct returns at least one value, and the *last* value is
		// is an error type, assume that is where error should go. This supports function types
		// that return just error, and functions that return (T, error) where T is some other non-error type
		//
		// Otherwise, if the template function has no return value, and an error occurred, panic with the error
		if functype.NumOut() > 0 && isErrorType(functype.Out(functype.NumOut()-1)) {
			results = append(results, out[1])
		} else {
			if err != nil {
				panic(err)
			}
		}

		return results
	}
}

func isErrorType(t reflect.Type) bool {
	errtype := reflect.TypeOf((*error)(nil)).Elem()
	return t.Implements(errtype)
}
