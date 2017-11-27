package ole1c

import "time"

type Constant interface {
}

type ConstantString struct {
	Get func() string `ole:"method"`
	Set func(value interface{}) `ole:"method"`
}

type ConstantTime struct {
	Get func() time.Time `ole:"method"`
	Set func(value interface{}) `ole:"method"`
}
