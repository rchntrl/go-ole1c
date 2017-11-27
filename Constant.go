package ole1c

type Constant struct {
	Get func() interface{} `ole:"method"`
	Set func(value interface{}) `ole:"method"`
}
