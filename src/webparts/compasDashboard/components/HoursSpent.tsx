import * as React from "react";
import classes from "./HoursSpent.module.scss";
import { Add } from "@material-ui/icons";
import { TextField, InputLabel } from "@material-ui/core";
import { DatePicker, IconButton, IIconProps } from "office-ui-fabric-react";
import { useState, useEffect } from "react";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { Done } from "@material-ui/icons";

const moreIcon = require("../../../ExternalRef/img/more.png");
const clearIcon: IIconProps = { iconName: "Clear" };
let latestId;
let EditId;

const InitialTime = {
  hours: "",
  Date: null,
  comName: "",
};

let dropValue = [];

let edit = [];
let setdata = [];

const HoursSpent = (props: any) => {
  const [hoursSec, setHoursSec] = useState(0);
  const [addNewAATLO, setAddNewAATLO] = useState({
    hours: "",
    Date: null,
    comName: "",
  });
  const [addNewJJ, setAddNewJJ] = useState({
    hours: "",
    Date: null,
    comName: "",
  });
  const [editHour, setEditHour] = useState(0);
  const [editValue, setEditValue] = useState(false);
  const [hoursSpentArr, setHoursSpentArr] = useState([]);
  const [addData, setAddData] = useState(false);
  const [addAALTO, setAddAALTO] = useState(0);
  const [addJOHN, setAddJOHN] = useState(0);
  const [totAdd, setTotAdd] = useState(0);
  const [listData, setListData] = useState({
    hours: "",
    Date: null,
    comName: "",
  });
  const [reEnter, setReEnter] = useState(true);
  const [show, setShow] = useState([false, 0]);
  // const [showOption, setShowOption] = useState(true);

  // Life Cycle of onload
  useEffect(() => {
    // setAddNewAATLO({
    //   hours: "",
    //   Date: null,
    //   comName: "",
    // });
    // setAddNewJJ({
    //   hours: "",
    //   Date: null,
    //   comName: "",
    // });

    latestId = 0;
    EditId = 0;
    if (props.Latest.key != 0 || props.Edit.item != null) {
      latestId = props.Latest.key;
      EditId = props.Edit.item;
      setHoursSec(latestId);
      setEditHour(EditId);
    } else {
      setHoursSec(latestId);
      setEditHour(EditId);
    }

    // Dropdown value taken
    props.sp.web.lists
      .getByTitle("Spent time")
      .fields.top(4000)
      .getByInternalNameOrTitle("CASCompany")()
      .then((response) => {
        console.log(response);
        dropValue = [];
        dropValue = response.Choices;
      })
      .then(() => {
        props.Edit.item != null
          ? props.sp.web.lists
              .getByTitle("Spent time")
              .items.top(4000)
              .select("*", "CASRef/ID")
              .filter(`CASRefId eq '${props.Edit.item}'`)
              .expand("CASRef")
              .orderBy("Modified", false)
              .get()
              .then((response) => {
                console.log(response);
                let hoursArr = [];
                hoursArr =
                  response.length > 0
                    ? response.map((res) => ({
                        Hours: res.CASHours ? res.CASHours : null,
                        Date: res.CASDate ? new Date(res.CASDate) : null,
                        Company: res.CASCompany,
                        spentId: res.Id,
                        isEdit: false,
                        showOption: false,
                      }))
                    : [];
                setHoursSpentArr([]);
                setHoursSpentArr([...hoursArr]);
                setAddData(true);
              })
              .catch((error) => {
                console.log(error);
              })
          : [];
      })
      .catch((error) => {
        console.log(error);
      });
    setReEnter(false);
  }, [reEnter]);

  useEffect(() => {
    let AALTOAdd = 0;
    let JOHNSONAdd = 0;
    let Total = 0;
    setAddAALTO(0);
    setAddJOHN(0);
    setTotAdd(0);
    hoursSpentArr.map((e) => {
      if (e.Company == "AALTO") {
        AALTOAdd = AALTOAdd + e.Hours;
        setAddAALTO(AALTOAdd);
      } else {
        JOHNSONAdd = JOHNSONAdd + e.Hours;
        setAddJOHN(JOHNSONAdd);
      }
      Total = AALTOAdd + JOHNSONAdd;
      setTotAdd(Total);
    });
    setAddData(false);
  }, [addData]);

  // GetDatas section
  const GetDatas = (company) => {
    let isValue =
      // addNewAATLO.comName == "AALTO" ? ValitationAA() : ValitationJJ();
      company == "Aalto" ? ValitationAA() : ValitationJJ();
    if (isValue) {
      AddHours(company);
    }
  };

  // Valition on AALTO Section
  const ValitationAA = () => {
    let isCheck = true;
    if (!addNewAATLO.hours) {
      isCheck = false;
      alertify.error("Please add Hour");
    } else if (!addNewAATLO.Date) {
      isCheck = false;
      alertify.error("Please add Date");
    }
    return isCheck;
  };

  // Valition on JOHNSON & JOHNSON Section
  const ValitationJJ = () => {
    let isCheck = true;
    if (!addNewJJ.hours) {
      isCheck = false;
      alertify.error("Please add Hour");
    } else if (!addNewJJ.Date) {
      isCheck = false;
      alertify.error("Please add Date");
    }
    return isCheck;
  };

  // Hours Add function
  const AddHours = (company) => {
    // dropValue.map((data) => {
    // return
    // addNewAATLO.comName == data
    company == "Aalto"
      ? props.sp.web.lists
          .getByTitle("Spent time")
          .items.add({
            Title:
              hoursSec != 0
                ? props.Latest.text
                : editHour == 0
                ? true
                : props.Edit.Title,
            CASHours: addNewAATLO.hours ? addNewAATLO.hours : null,
            CASDate: addNewAATLO.Date ? addNewAATLO.Date : null,
            CASCompany: addNewAATLO.comName ? addNewAATLO.comName : "",
            CASRefId:
              hoursSec != 0
                ? props.Latest.key
                : editHour == 0
                ? true
                : props.Edit.item,
          })
          .then((response) => {
            console.log(response);
            alertify.success("Record submitted successfully");
            addNewAATLO.hours = "";
            addNewAATLO.Date = null;
            setAddNewAATLO({
              ...addNewAATLO,
            });
            props.renderProject();
            setReEnter(true);
          })
      : // addNewJJ.comName == data &&
        props.sp.web.lists
          .getByTitle("Spent time")
          .items.add({
            Title:
              hoursSec != 0
                ? props.Latest.text
                : editHour == 0
                ? true
                : props.Edit.Title,
            CASHours: addNewJJ.hours ? addNewJJ.hours : null,
            CASDate: addNewJJ.Date ? addNewJJ.Date : null,
            CASCompany: addNewJJ.comName ? addNewJJ.comName : "",
            CASRefId:
              hoursSec != 0
                ? props.Latest.key
                : editHour == 0
                ? true
                : props.Edit.item,
          })
          .then((response) => {
            console.log(response);
            alertify.success("Record submitted successfully");
            addNewJJ.hours = "";
            addNewJJ.Date = null;
            setAddNewJJ({
              ...addNewJJ,
            });
            props.renderProject();
            setReEnter(true);
          });
    // });
  };

  // Save Data function
  const SaveData = (data) => {
    let isAllFilled = false;
    listData.hours == "" || listData.hours == "0" || !listData.Date
      ? (isAllFilled = false)
      : (isAllFilled = true);
    isAllFilled
      ? props.sp.web.lists
          .getByTitle("Spent time")
          .items.getById(data)
          .update({
            CASHours: listData.hours,
            CASDate: listData.Date,
          })
          .then((e) => {
            console.log(e);
            alertify.success("Record submitted successfully");
            props.renderProject();
            setReEnter(true);
          })
          .catch((e) => {
            console.log(e);
          })
      : listData.hours == "" || listData.hours == "0"
      ? alertify.error("Please add Hour")
      : alertify.error("Please add Date");
  };

  // Delete function
  const getDelete = (data) => {
    props.sp.web.lists
      .getByTitle("Spent time")
      .items.getById(data)
      .delete()
      .then((e) => {
        console.log(e);
        props.renderProject();
        setReEnter(true);
      })
      .catch((err) => {
        console.log(err);
      });
  };

  return (
    <>
      <div className={classes.contentTop}>
        <div className={classes["c1"]}>
          <div
            className={classes.title3}
            style={{ transform: "translateY(-6px)" }}
          >
            AALTO
          </div>

          <div className={classes.inputs}>
            {/* AALTO Company section hr */}
            <div className={classes.inpField}>
              <InputLabel className={classes.inpLabel}>Hours Spent:</InputLabel>
              <TextField
                disabled={hoursSec != 0 ? false : editHour == 0 ? true : false}
                className={`${classes.inpt3} HSTextField`}
                id="standard-basic"
                variant="outlined"
                placeholder={`0,0h`}
                type="number"
                InputLabelProps={{ shrink: false }}
                value={addNewAATLO.hours}
                onChange={(e) => {
                  addNewAATLO.hours =
                    +e.target.value < 0 ? "0" : e.target.value;
                  addNewAATLO.comName = "AALTO";
                  setAddNewAATLO({ ...addNewAATLO });
                }}
              />
            </div>

            {/* AALTO Company section Date */}
            <div className={classes.inpField}>
              <InputLabel className={classes.inpLabel}>yyyy/mm:</InputLabel>
              <DatePicker
                disabled={hoursSec != 0 ? false : editHour == 0 ? true : false}
                className={`HSDatePicker ${classes.datet3}`}
                placeholder={`0000/00`}
                formatDate={(date: Date): string => {
                  return `${date.toLocaleDateString().split("/")[2]}/${
                    +date.toLocaleDateString().split("/")[0] < 10
                      ? "0" + date.toLocaleDateString().split("/")[0]
                      : date.toLocaleDateString().split("/")[0]
                  }`;
                }}
                textField={{
                  onRenderSuffix: () => (
                    <IconButton
                      onClick={() => {
                        addNewAATLO.Date = "";
                        setAddNewAATLO({ ...addNewAATLO });
                      }}
                      iconProps={clearIcon}
                    />
                  ),
                  styles: { suffix: { padding: "0 4px" } },
                }}
                value={addNewAATLO.Date ? new Date(addNewAATLO.Date) : null}
                onSelectDate={(e) => {
                  addNewAATLO.Date = e.toISOString();
                  setAddNewAATLO({ ...addNewAATLO });
                }}
              />
            </div>

            {/* AALTO Company Button Section */}
            <button
              disabled={hoursSec != 0 ? false : editHour == 0 ? true : false}
              className={classes.AddBtn}
              onClick={() => GetDatas("Aalto")}
            >
              <Add
                style={{
                  backgroundColor: "transparent",
                  borderRadius: "50%",
                  width: "40px",
                  border: "1px solid",
                  padding: "0px 6px",
                  marginLeft: "10px",
                  marginTop: "40px",
                  transform: "scale(1.2)",
                }}
              />
            </button>
          </div>
        </div>
        <div className={classes["c1"]}>
          <div className={classes.title3}>JOHNSON & JOHNSON</div>

          <div className={classes.inputs}>
            {/* JOHNSON & JOHNSON Company section hr */}
            <div className={classes.inpField}>
              <InputLabel className={classes.inpLabel}>Hours Spent:</InputLabel>
              <TextField
                disabled={hoursSec != 0 ? false : editHour == 0 ? true : false}
                className={`${classes.inpt3} HSTextField`}
                id="standard-basic"
                variant="outlined"
                placeholder={`0,0h`}
                type="number"
                InputLabelProps={{ shrink: false }}
                value={addNewJJ.hours}
                onChange={(e) => {
                  addNewJJ.hours = +e.target.value < 0 ? "0" : e.target.value;
                  addNewJJ.comName = "JOHNSON and JOHNSON";
                  setAddNewJJ({ ...addNewJJ });
                }}
              />
            </div>

            {/* JOHNSON & JOHNSON Company section Date */}
            <div className={classes.inpField}>
              <InputLabel className={classes.inpLabel}>yyyy/mm:</InputLabel>
              <DatePicker
                disabled={hoursSec != 0 ? false : editHour == 0 ? true : false}
                className={`${classes.datet3} HSDatePicker`}
                placeholder={`0000/00`}
                formatDate={(date: Date): string => {
                  return `${date.toLocaleDateString().split("/")[2]}/${
                    +date.toLocaleDateString().split("/")[0] < 10
                      ? "0" + date.toLocaleDateString().split("/")[0]
                      : date.toLocaleDateString().split("/")[0]
                  }`;
                }}
                value={addNewJJ.Date ? new Date(addNewJJ.Date) : null}
                onSelectDate={(e) => {
                  addNewJJ.Date = e.toISOString();
                  setAddNewJJ({ ...addNewJJ });
                }}
                textField={{
                  onRenderSuffix: () => (
                    <IconButton
                      onClick={() => {
                        addNewJJ.Date = "";
                        setAddNewJJ({ ...addNewJJ });
                      }}
                      iconProps={clearIcon}
                    />
                  ),
                  styles: { suffix: { padding: "0 4px" } },
                }}
              />
            </div>

            {/* JOHNSON & JOHNSON Company Button Section */}
            <button
              disabled={hoursSec != 0 ? false : editHour == 0 ? true : false}
              className={classes.AddBtn}
              onClick={() => GetDatas("j&j")}
            >
              <Add
                style={{
                  backgroundColor: "transparent",
                  borderRadius: "50%",
                  width: "40px",
                  border: "1px solid",
                  padding: "0px 6px",
                  marginLeft: "10px",
                  marginTop: "40px",
                  transform: "scale(1.2)",
                }}
              />
            </button>
          </div>
        </div>{" "}
      </div>

      <div className={classes.contentBottom}>
        <div className={classes.outputs}>
          {/* AALTO Company Section */}
          <div className={`${classes.o1} optionOne`}>
            {hoursSpentArr.map((e) => {
              return (
                <>
                  {e.Company == "AALTO" && (
                    <div className={classes.inpField}>
                      <div>
                        {e.isEdit ? (
                          <TextField
                            type="number"
                            disabled={!e.isEdit}
                            className={classes.inpt3}
                            value={+listData.hours}
                            onChange={(e) => {
                              listData.hours =
                                +e.target.value < 0 ? "0" : e.target.value;
                              setListData({ ...listData });
                            }}
                          />
                        ) : (
                          <InputLabel
                            className={`${classes.hourView} hourViewOne`}
                            style={{ width: "172px" }}
                          >
                            {`${String(e.Hours).split(".").join(",")} `}{" "}
                            <span className={classes.hourViewHour}>h</span>
                          </InputLabel>
                        )}
                      </div>
                      <div>
                        {e.isEdit ? (
                          <DatePicker
                            disabled={!e.isEdit}
                            placeholder={`0000/00`}
                            className={`${classes.datet3} HSEditDatePicker`}
                            formatDate={(date: Date): string => {
                              return `${
                                date.toLocaleDateString().split("/")[2]
                              }/${
                                +date.toLocaleDateString().split("/")[0] < 10
                                  ? "0" +
                                    date.toLocaleDateString().split("/")[0]
                                  : date.toLocaleDateString().split("/")[0]
                              }`;
                            }}
                            onSelectDate={(date) => {
                              listData.Date = date.toISOString();
                              setListData({ ...listData });
                            }}
                            textField={{
                              onRenderSuffix: () => (
                                <IconButton
                                  onClick={() => {
                                    listData.Date = null;
                                    setListData({ ...listData });
                                  }}
                                  iconProps={clearIcon}
                                />
                              ),
                              styles: { suffix: { padding: "0 4px" } },
                            }}
                            value={
                              listData.Date ? new Date(listData.Date) : null
                            }
                          />
                        ) : (
                          <InputLabel
                            className={classes.dateView}
                            style={{ width: 162 }}
                          >
                            {`${e.Date.toLocaleDateString().split("/")[2]}/${
                              +e.Date.toLocaleDateString().split("/")[0] < 10
                                ? "0" +
                                  e.Date.toLocaleDateString().split("/")[0]
                                : e.Date.toLocaleDateString().split("/")[0]
                            }`}
                          </InputLabel>
                        )}
                      </div>
                      <div
                        className={classes.options}
                        style={{ marginLeft: 20 }}
                      >
                        {!e.isEdit ? (
                          <>
                            {hoursSpentArr.every(
                              (data) => data.isEdit == false
                            ) && (
                              <img
                                style={{ cursor: "pointer" }}
                                src={`${moreIcon}`}
                                width={18}
                                height={20}
                                alt="more"
                                onClick={() => {
                                  hoursSpentArr.forEach((row) => {
                                    row.showOption = false;
                                  });
                                  hoursSpentArr.filter(
                                    (row) => row.spentId == e.spentId
                                  )[0].showOption = true;
                                  setHoursSpentArr([...hoursSpentArr]);
                                }}
                              />
                            )}
                            {e.showOption ? (
                              <div className={classes.optionSection}>
                                <div
                                  style={{
                                    borderBottom: "1px solid #cacaca",
                                  }}
                                  onClick={() => {
                                    hoursSpentArr.forEach((row) => {
                                      row.showOption = false;
                                    });
                                    hoursSpentArr.filter(
                                      (row) => row.spentId == e.spentId
                                    )[0].isEdit = true;
                                    setHoursSpentArr([...hoursSpentArr]);
                                    listData.Date = e.Date;
                                    listData.hours = e.Hours;
                                    setListData({ ...listData });
                                  }}
                                >
                                  Edit
                                </div>
                                <div onClick={() => getDelete(e.spentId)}>
                                  Cancel
                                </div>
                              </div>
                            ) : (
                              ""
                            )}
                          </>
                        ) : (
                          <Done
                            className={classes.doneIcon}
                            style={{
                              cursor: "pointer",
                              color: "green",
                            }}
                            onClick={() => {
                              SaveData(e.spentId);
                            }}
                          />
                        )}
                      </div>
                    </div>
                  )}
                </>
              );
            })}
            <div>
              <div></div>
              <div></div>
            </div>
          </div>

          {/* JOHNSON & JOHNSON Company Section */}
          <div
            className={`${classes.o1} ${classes.o1_1}`}
            style={{ paddingLeft: "16px" }}
          >
            {hoursSpentArr.map((e) => {
              return (
                <>
                  {e.Company.toLowerCase() == "johnson and johnson" && (
                    <div className={classes.inpField}>
                      <div>
                        {e.isEdit ? (
                          <TextField
                            type="number"
                            disabled={!e.isEdit}
                            className={classes.inpt3}
                            value={+listData.hours}
                            onChange={(e) => {
                              listData.hours =
                                +e.target.value < 0 ? "0" : e.target.value;
                              setListData({ ...listData });
                            }}
                          />
                        ) : (
                          <InputLabel
                            className={classes.hourView}
                            style={{ width: 170 }}
                          >
                            {`${String(e.Hours).split(".").join(",")} `}{" "}
                            <span className={classes.hourViewHour}>h</span>
                          </InputLabel>
                        )}
                      </div>
                      <div>
                        {e.isEdit ? (
                          <DatePicker
                            disabled={!e.isEdit}
                            placeholder={`0000/00`}
                            className={`${classes.datet3} HSEditDatePicker`}
                            formatDate={(date: Date): string => {
                              return `${
                                date.toLocaleDateString().split("/")[2]
                              }/${
                                +date.toLocaleDateString().split("/")[0] < 10
                                  ? "0" +
                                    date.toLocaleDateString().split("/")[0]
                                  : date.toLocaleDateString().split("/")[0]
                              }`;
                            }}
                            onSelectDate={(date) => {
                              listData.Date = date.toISOString();
                              setListData({ ...listData });
                            }}
                            textField={{
                              onRenderSuffix: () => (
                                <IconButton
                                  onClick={() => {
                                    listData.Date = null;
                                    setListData({ ...listData });
                                  }}
                                  iconProps={clearIcon}
                                />
                              ),
                              styles: { suffix: { padding: "0 4px" } },
                            }}
                            value={
                              listData.Date ? new Date(listData.Date) : null
                            }
                          />
                        ) : (
                          <InputLabel
                            className={classes.dateView}
                            style={{ width: 170 }}
                          >
                            {`${e.Date.toLocaleDateString().split("/")[2]}/${
                              +e.Date.toLocaleDateString().split("/")[0] < 10
                                ? "0" +
                                  e.Date.toLocaleDateString().split("/")[0]
                                : e.Date.toLocaleDateString().split("/")[0]
                            }`}
                          </InputLabel>
                        )}
                      </div>
                      <div className={classes.options}>
                        {!e.isEdit ? (
                          <>
                            {hoursSpentArr.every(
                              (data) => data.isEdit == false
                            ) && (
                              <img
                                style={{ cursor: "pointer" }}
                                src={`${moreIcon}`}
                                width={18}
                                height={20}
                                alt="more"
                                onClick={() => {
                                  hoursSpentArr.forEach((row) => {
                                    row.showOption = false;
                                  });
                                  hoursSpentArr.filter(
                                    (row) => row.spentId == e.spentId
                                  )[0].showOption = true;
                                  setHoursSpentArr([...hoursSpentArr]);
                                }}
                              />
                            )}
                            {e.showOption ? (
                              <div className={classes.optionSection}>
                                <div
                                  style={{
                                    borderBottom: "1px solid #cacaca",
                                  }}
                                  onClick={() => {
                                    hoursSpentArr.forEach((row) => {
                                      row.showOption = false;
                                    });
                                    hoursSpentArr.filter(
                                      (row) => row.spentId == e.spentId
                                    )[0].isEdit = true;
                                    setHoursSpentArr([...hoursSpentArr]);
                                    listData.Date = e.Date;
                                    listData.hours = e.Hours;
                                    setListData({ ...listData });
                                  }}
                                >
                                  Edit
                                </div>
                                <div onClick={() => getDelete(e.spentId)}>
                                  Cancel
                                </div>
                              </div>
                            ) : (
                              ""
                            )}
                          </>
                        ) : (
                          <Done
                            className={classes.doneIcon}
                            style={{
                              cursor: "pointer",
                              color: "green",
                            }}
                            onClick={() => {
                              SaveData(e.spentId);
                            }}
                          />
                        )}
                      </div>
                    </div>
                  )}
                </>
              );
            })}
            <div className={classes.totalSection}>
              <div></div>
              <div></div>
            </div>
          </div>
        </div>
        <div className={classes.totals}>
          <div className={classes.total}>
            <div>{`${String(addAALTO).split(".")[0]},${
              String(addAALTO).split(".").length == 1
                ? "00"
                : +String(addAALTO).split(".")[1] > 9
                ? String(addAALTO).split(".")[1]
                : `${String(addAALTO).split(".")[1]}0`
            }`}</div>
            <div>AAlto tot.h </div>
          </div>
          <div className={classes.total}>
            <div>
              {`${String(addJOHN).split(".")[0]},${
                String(addJOHN).split(".").length == 1
                  ? "00"
                  : +String(addJOHN).split(".")[1] > 9
                  ? String(addJOHN).split(".")[1]
                  : `${String(addJOHN).split(".")[1]}0`
              }`}
            </div>
            <div>J&J tot.h </div>
          </div>
        </div>
        <div className={classes.outputs} style={{ border: "none" }}>
          <div className={classes.o2}>
            {/* {String(totAdd).split(".").join(",")} */}
            {`${String(totAdd).split(".")[0]},${
              String(totAdd).split(".").length == 1
                ? "00"
                : +String(totAdd).split(".")[1] > 9
                ? String(totAdd).split(".")[1]
                : `${String(totAdd).split(".")[1]}0`
            }`}
            <span>tot.h </span>
          </div>
        </div>
      </div>
    </>
  );
};

export default HoursSpent;
