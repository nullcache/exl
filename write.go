// Copyright 2022 exl Author. All Rights Reserved.
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//      http://www.apache.org/licenses/LICENSE-2.0
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package exl

import (
	"fmt"
	"github.com/tealeg/xlsx/v3"
	"io"
	"reflect"
	"strings"
	"time"
)

type (
	WriteConfigurator interface{ WriteConfigure(wc *WriteConfig) }
	WriteConfig       struct {
		SheetName string
		TagName   string
		// Skip when struct field have NOT matched tagName.
		SkipNoTag bool
		// Skip when struct field is a nil pointer.
		SkipNilPointer bool
		// Set dropList and write value which is transformed from key.
		DropListMap map[string][]struct {
			Key   string
			Value string
		}
		// Transform TRUE/FALSE to Chinese 是/否.
		ChineseBool  bool
		WriteTimeFmt string
	}
)

var defaultWriteConfig = func() *WriteConfig {
	return &WriteConfig{SheetName: "Sheet1", TagName: "excel", WriteTimeFmt: xlsx.DefaultDateFormat}
}

func write(sheet *xlsx.Sheet, data []any, wc ...*WriteConfig) {
	var wConfig *WriteConfig
	if len(wc) >= 0 {
		wConfig = wc[0]
	}
	r := sheet.AddRow()
	for _, cell := range data {
		if reflect.TypeOf(cell) == reflect.TypeOf(time.Time{}) {
			r.AddCell().SetDateWithOptions(cell.(time.Time), xlsx.DateTimeOptions{
				Location:        xlsx.DefaultDateOptions.Location,
				ExcelTimeFormat: wConfig.WriteTimeFmt,
			})
		} else {
			r.AddCell().SetValue(cell)
		}
	}
}

func NewFileFromSlice[T WriteConfigurator](ts []T) *xlsx.File {
	f := xlsx.NewFile()
	write0(f, ts)
	return f
}

// WriteFile defines write []T to excel file
//
// params: file,excel file full path
//
// params: typed parameter T, must be implements exl.Bind
func WriteFile[T WriteConfigurator](file string, ts []T) error {
	f := xlsx.NewFile()
	write0(f, ts)
	return f.Save(file)
}

// WriteTo defines write to []T to excel file
//
// params: w, the dist writer
//
// params: typed parameter T, must be implements exl.Bind
func WriteTo[T WriteConfigurator](w io.Writer, ts []T) error {
	f := xlsx.NewFile()
	write0(f, ts)
	return f.Write(w)
}

func write0[T WriteConfigurator](f *xlsx.File, ts []T) {
	wc := defaultWriteConfig()
	if len(ts) > 0 {
		ts[0].WriteConfigure(wc)
	}
	haveDropList := wc.DropListMap != nil

	tT := new(T)
	if sheet, _ := f.AddSheet(wc.SheetName); sheet != nil {
		typ := reflect.TypeOf(tT).Elem().Elem()
		numField := typ.NumField()
		header := make([]any, 0, numField)
		for i := 0; i < numField; i++ {
			fe := typ.Field(i)
			if !fe.IsExported() {
				continue
			}
			name := fe.Name
			tt, have := fe.Tag.Lookup(wc.TagName)
			if have {
				name = tt
			}
			if have || !wc.SkipNoTag {
				header = append(header, name)
			}
		}
		// write header
		write(sheet, header, wc)
		if len(ts) > 0 {
			// write data
			for i1, t := range ts {
				data := make([]any, 0, numField)
				for i := 0; i < numField; i++ {
					rowIndex := i1 + 1
					colIndex := len(data)

					v := reflect.ValueOf(t).Elem().Field(i)
					if !v.CanInterface() {
						continue
					}
					tag, have := reflect.TypeOf(t).Elem().Field(i).Tag.Lookup(wc.TagName)
					if !have && wc.SkipNoTag {
						continue
					}

					// 1. add validation
					basicType := v.Kind()
					if v.Kind() == reflect.Ptr {
						basicType = v.Type().Elem().Kind()
					}

					if basicType == reflect.Bool {
						dd := xlsx.NewDataValidation(rowIndex, colIndex, rowIndex, colIndex, v.Kind() == reflect.Ptr)
						if wc.ChineseBool {
							dd.SetDropList([]string{"是", "否"})
							errTitle := ""
							errMsg := "应该为 是或否"
							dd.SetError(xlsx.StyleStop, &errTitle, &errMsg)
							sheet.AddDataValidation(dd)
						} else {
							dd.SetDropList([]string{"TRUE", "FALSE"})
							errTitle := ""
							errMsg := "should be TRUE or FALSE"
							dd.SetError(xlsx.StyleStop, &errTitle, &errMsg)
							sheet.AddDataValidation(dd)
						}
					}

					if basicType == reflect.String {
						if haveDropList {
							dropList, have := wc.DropListMap[tag]
							if have {
								dd := xlsx.NewDataValidation(rowIndex, colIndex, rowIndex, colIndex, v.Kind() == reflect.Ptr)
								dropListArr := make([]string, 0, len(dropList))
								for _, v := range dropList {
									dropListArr = append(dropListArr, v.Value)
								}
								dd.SetDropList(dropListArr)
								errTitle := ""
								errMsg := fmt.Sprintf("应该为 %s 中之一", strings.Join(dropListArr, "、"))
								dd.SetError(xlsx.StyleStop, &errTitle, &errMsg)
								sheet.AddDataValidation(dd)
							}
						}
					}

					// 2. add special data
					if v.Kind() == reflect.Ptr {
						if wc.SkipNilPointer && v.IsNil() {
							data = append(data, "")
							continue
						} else if !v.IsNil() {
							v = v.Elem()
						}
					}
					if v.Kind() == reflect.Bool {
						if wc.ChineseBool {
							if v.Bool() {
								data = append(data, interface{}("是"))
							} else {
								data = append(data, interface{}("否"))
							}
						} else {
							data = append(data, v.Interface())
						}
						continue
					}

					if v.Kind() == reflect.String {
						if haveDropList {
							dropList, have := wc.DropListMap[tag]
							if have {
								key := v.String()
								value := key
								for _, v := range dropList {
									if v.Key == key {
										value = v.Value
									}
								}
								data = append(data, interface{}(value))
								continue
							}
						}
					}
					data = append(data, v.Interface())

				}
				write(sheet, data, wc)
			}
		}
	}
}

// WriteExcel defines write [][]string to excel
//
// params: file, excel file pull path
//
// params: data, write data to excel
func WriteExcel(file string, data [][]string) error {
	f := xlsx.NewFile()
	writeExcel0(f, data)
	return f.Save(file)
}

// WriteExcelTo defines write [][]string to excel
//
// params: w, the dist writer
//
// params: data, write data to excel
func WriteExcelTo(w io.Writer, data [][]string) error {
	f := xlsx.NewFile()
	writeExcel0(f, data)
	return f.Write(w)
}

func writeExcel0(f *xlsx.File, data [][]string) {
	sheet, _ := f.AddSheet("Sheet1")
	for _, row := range data {
		r := sheet.AddRow()
		for _, cell := range row {
			r.AddCell().SetString(cell)
		}
	}
}
