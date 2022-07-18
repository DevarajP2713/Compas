import * as React from "react";
import { useState, useEffect, useRef } from "react";
import classes from "./Panel.module.scss";
import {
  TextField,
  Typography,
  Box,
  Tab,
  Tabs,
  InputLabel,
  MenuItem,
  Select,
} from "@material-ui/core";
// import PropTypes from "prop-types";
import * as PropTypes from "prop-types";
import {
  Add,
  Alarm,
  CalendarToday,
  Close,
  Comment,
  Edit,
  Mail,
  ModeComment,
  Send,
} from "@material-ui/icons";
import {
  DatePicker,
  ILabelStyles,
  IStyleSet,
  Label,
  Pivot,
  PivotItem,
} from "office-ui-fabric-react";
import { initial } from "lodash";
import { log } from "@pnp/pnpjs";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import LatestAction from "./LatestAction";
import HoursSpent from "./HoursSpent";

const ProjectIcon = require("../../../ExternalRef/img/ProjectEdit.png");
const LatestActionIcon = require("../../../ExternalRef/img/LatestAction.png");
const HoursSpentIcon = require("../../../ExternalRef/img/HoursSpent.png");

const objProjInfo = {
  ProjectName: "",
  Priority: null,
  CountryIBVT: null,
  OrganizationUnit: "",
  EngagemantType: null,
  EngagementSubType: null,
  Requestor: [],
  StatusType: null,
  IDNumber: null,
  CreationDate: null,
  LastModifyDate: null,
  EngagementScope: null,
  Actions: "",
  CrossCharge: "",
  ProjectStartDate: null,
  ProjectCompletionDate: null,
  EngagementNotes: "",
};

let RequestorIdArr = [];

let priorityArr = [];
let CountriesArr = [];
let EngagementTypeArr = [];
let EngagementTypeSubArr = [];
let RequestorArr = [];
let StatusTypeArr = [];
let EngagementScopeArr = [];
let ProjectArr = [];

const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
  root: { marginTop: 10 },
};

let objForAction = {
  ID: 0,
  Title: "",
};

const Panel = (props: any) => {
  const [value, setValue] = useState(0);
  // const [Priority, setPriority] = useState("PlaceHolder");
  const [addDatas, setAddDatas] = useState(objProjInfo);
  const [isEdit, setIsEdit] = useState(false);
  const [projectName, setProjectName] = useState("");
  const [priorityChoice, setPriorityChoice] = useState(priorityArr);
  const [countryChoice, setCountryChoice] = useState(CountriesArr);
  const [engTypeChoice, setEngTypeChoice] = useState(EngagementTypeArr);
  const [engSubTypeChoice, setEngSubTypeChoice] =
    useState(EngagementTypeSubArr);
  const [statusTypeChoice, setStatusTypeChoice] = useState(StatusTypeArr);
  const [engScopeChoice, setEngScopeChoice] = useState(EngagementScopeArr);
  const [isDataBind, setIsDataBind] = useState(false);
  const [btnDisabled, setBtnDisabled] = useState(true);
  const [peopleId, setPeopleId] = useState([]);
  const [latestSec, setLatestSec] = useState({ key: 0, text: "" });
  const [forAction, setForAction] = useState(objForAction);
  const [reRenderTable, setReRenderTable] = useState(false);
  const [selectedKey, setSelectedKey] = useState(+props.Edit.num);

  const handleChange = (event, newValue) => {
    setValue(newValue);
  };

  useEffect(() => {
    props.renderProject();
  }, [reRenderTable]);

  // Life Cycle of Onload
  useEffect(() => {
    setAddDatas({
      ProjectName: "",
      Priority: null,
      CountryIBVT: null,
      OrganizationUnit: "",
      EngagemantType: null,
      EngagementSubType: 0,
      Requestor: [],
      StatusType: null,
      IDNumber: null,
      CreationDate: null,
      LastModifyDate: null,
      EngagementScope: null,
      Actions: "",
      CrossCharge: "",
      ProjectStartDate: null,
      ProjectCompletionDate: null,
      EngagementNotes: "",
    });

    // Projects Id Taken
    props.sp.web.lists
      .getByTitle("Projects")
      .items.top(4000)
      .get()
      .then((response) => {
        ProjectArr = response.map((data) => data.ID);
      })
      .catch((error) => {
        console.log(error);
      });

    // Priority Values Taken
    props.sp.web.lists
      .getByTitle("Priorities")
      .items.top(4000)
      .get()
      .then((response) => {
        priorityArr = response.map((data) => ({
          key: data.ID,
          text: data.Title,
          icon: data.Icon,
        }));
        setPriorityChoice(priorityArr);
      })
      .catch((error) => {
        console.log(error);
      })
      .then(() => {
        // Country Values Taken
        props.sp.web.lists
          .getByTitle("Countries")
          .items.top(4000)
          .select("ID", "Title")
          .get()
          .then((response) => {
            CountriesArr = response.map((data) => ({
              key: data.ID,
              text: data.Title,
            }));
            CountriesArr.sort((a, b) => a.text.localeCompare(b.text));
            setCountryChoice(CountriesArr);
          })
          .then(() => {
            // Engagement Type Values Taken
            props.sp.web.lists
              .getByTitle("Engagement Types")
              .items.top(4000)
              .select("ID", "Title")
              .get()
              .then((response) => {
                EngagementTypeArr = response.map((data) => ({
                  key: data.ID,
                  text: data.Title,
                }));
                setEngTypeChoice(EngagementTypeArr);
              })
              .catch((error) => {
                console.log(error);
              })
              .then(() => {
                // Status Type Values Taken
                props.sp.web.lists
                  .getByTitle("Status types")
                  .items.top(4000)
                  .select("ID", "Title")
                  .get()
                  .then((response) => {
                    StatusTypeArr = response.map((data) => ({
                      key: data.ID,
                      text: data.Title,
                    }));
                    setStatusTypeChoice(StatusTypeArr);
                  })
                  .then(() => {
                    // Engagement Scope Values Taken
                    props.sp.web.lists
                      .getByTitle("Engagement Scopes")
                      .items.top(4000)
                      .select("ID", "Title")
                      .get()
                      .then((response) => {
                        EngagementScopeArr = response.map((data) => {
                          return {
                            key: data.ID,
                            text: data.Title,
                          };
                        });
                        setEngScopeChoice(EngagementScopeArr);
                      })
                      .catch((error) => {
                        console.log(error);
                      });
                  })
                  .then(() => {
                    props.sp.web.lists
                      .getByTitle("Engagement subtypes")
                      .items.top(4000)
                      .select("*", "CASEngType/Title", "CASEngType/ID")
                      .expand("CASEngType")
                      .get()
                      .then((res) => {
                        EngagementTypeSubArr = res.map((rs) => ({
                          key: rs.ID,
                          text: rs.Title,
                          type: rs.CASEngType.ID,
                        }));
                      });
                  })
                  .catch((error) => {
                    console.log(error);
                  });
              });
          })
          .catch((error) => {
            console.log(error);
          });
      });
    setIsEdit(false);
    setIsDataBind(true);
  }, []);

  useEffect(() => {
    props.sp.web.lists
      .getByTitle("Engagement subtypes")
      .items.top(4000)
      .select("*", "CASEngType/Title", "CASEngType/ID")
      .expand("CASEngType")
      .get()
      .then((res) => {
        EngagementTypeSubArr = res.map((rs) => ({
          key: rs.ID,
          text: rs.Title,
          type: rs.CASEngType.ID,
        }));
      });
    props.Edit.flagEdit
      ? props.sp.web.lists
          .getByTitle("Projects")
          .items.top(4000)
          .getById(props.Edit.item)
          .select("*", "CASUser/EMail", "CASUser/Id")
          .expand("CASUser")
          .get()
          .then((response) => {
            console.log(response);
            let PeoEMailArr = [];

            if (response.CASUser) {
              setPeopleId(response.CASUser.map((res) => res.Id));
              PeoEMailArr = response.CASUser.map((res) => res.EMail);
            }
            setEngSubTypeChoice(
              response.CASEngTypeId
                ? [
                    ...EngagementTypeSubArr.filter(
                      (choice) => choice.type == response.CASEngTypeId
                    ),
                  ]
                : []
            );
            setAddDatas({
              ProjectName: response.Title ? response.Title : "",
              Priority: response.CASPriorityId ? response.CASPriorityId : 0,
              CountryIBVT: response.CASCountryId ? response.CASCountryId : 0,
              OrganizationUnit: response.CASOrgUnit ? response.CASOrgUnit : "",
              EngagemantType: response.CASEngTypeId ? response.CASEngTypeId : 0,
              EngagementSubType: response.CASEngSubTypeId
                ? response.CASEngSubTypeId
                : 0,
              Requestor: PeoEMailArr ? PeoEMailArr : [],
              StatusType: response.CASStatusId ? response.CASStatusId : 0,
              IDNumber: response.ID,
              CreationDate: response.Created
                ? new Date(response.Created)
                : null,
              LastModifyDate: response.Modified
                ? new Date(response.Modified)
                : null,
              EngagementScope: response.CASEngScopeId
                ? response.CASEngScopeId
                : 0,
              Actions: "",
              CrossCharge: response.CASCCI ? response.CASCCI : "",
              ProjectStartDate: response.CASPrjStartDate
                ? new Date(response.CASPrjStartDate)
                : null,
              ProjectCompletionDate: response.CASPrjEndDate
                ? new Date(response.CASPrjEndDate)
                : null,
              EngagementNotes: response.CASEngNotes ? response.CASEngNotes : "",
            });
            setEngSubTypeChoice(
              response.CASEngTypeId
                ? [
                    ...EngagementTypeSubArr.filter(
                      (choice) => choice.type == response.CASEngTypeId
                    ),
                  ]
                : []
            );
            setBtnDisabled(false);
            setIsEdit(true);
          })
          .catch((error) => {
            console.log(error);
          })
      : setAddDatas({
          ProjectName: "",
          Priority: null,
          CountryIBVT: null,
          OrganizationUnit: "",
          EngagemantType: null,
          EngagementSubType: 0,
          Requestor: [],
          StatusType: null,
          IDNumber: null,
          CreationDate: null,
          LastModifyDate: null,
          EngagementScope: null,
          Actions: "",
          CrossCharge: "",
          ProjectStartDate: null,
          ProjectCompletionDate: null,
          EngagementNotes: "",
        });
  }, [isDataBind]);

  // Add Datas
  const AddListItems = () => {
    props.sp.web.lists
      .getByTitle("Projects")
      .items.add({
        Title: addDatas.ProjectName ? addDatas.ProjectName : "",
        CASPriorityId:
          addDatas.Priority && addDatas.Priority != "null"
            ? addDatas.Priority
            : 0,
        CASCountryId:
          addDatas.CountryIBVT && addDatas.CountryIBVT != "null"
            ? addDatas.CountryIBVT
            : 0,
        CASOrgUnit: addDatas.OrganizationUnit ? addDatas.OrganizationUnit : "",
        CASEngTypeId:
          addDatas.EngagemantType && addDatas.EngagemantType != "null"
            ? addDatas.EngagemantType
            : 0,
        CASEngSubTypeId: addDatas.EngagementSubType
          ? addDatas.EngagementSubType
          : 0,
        CASUserId: RequestorIdArr
          ? { results: RequestorIdArr }
          : { results: [] },
        CASStatusId:
          addDatas.StatusType && addDatas.StatusType != "null"
            ? addDatas.StatusType
            : 0,
        CASEngScopeId:
          addDatas.EngagementScope && addDatas.EngagementScope != "null"
            ? addDatas.EngagementScope
            : 0,
        CASCCI: addDatas.CrossCharge ? addDatas.CrossCharge : "",
        CASPrjStartDate: addDatas.ProjectStartDate
          ? addDatas.ProjectStartDate
          : null,
        CASPrjEndDate: addDatas.ProjectCompletionDate
          ? addDatas.ProjectCompletionDate
          : null,
        CASEngNotes: addDatas.EngagementNotes ? addDatas.EngagementNotes : "",
      })
      .then((response) => {
        objForAction = {
          ID: response.ID,
          Title: response.Title,
        };
        RequestorIdArr = [];
        setForAction({ ...objForAction });
        setLatestSec({ key: response.data.ID, text: response.data.Title });
        setAddDatas({
          ProjectName: "",
          Priority: null,
          CountryIBVT: null,
          OrganizationUnit: "",
          EngagemantType: null,
          EngagementSubType: 0,
          Requestor: [],
          StatusType: null,
          IDNumber: null,
          CreationDate: null,
          LastModifyDate: null,
          EngagementScope: null,
          Actions: "",
          CrossCharge: "",
          ProjectStartDate: null,
          ProjectCompletionDate: null,
          EngagementNotes: "",
        });
        alertify.success("Record submitted successfully");
        setReRenderTable(true);
        props.Panel(false);
      })
      .catch((error) => {
        console.log(error);
      });
  };

  // Update datas
  const UpdateListItem = () => {
    props.sp.web.lists
      .getByTitle("Projects")
      .items.getById(props.Edit.item)
      .update({
        Title: addDatas.ProjectName ? addDatas.ProjectName : "",
        CASPriorityId: addDatas.Priority ? addDatas.Priority : 0,
        CASCountryId: addDatas.CountryIBVT ? addDatas.CountryIBVT : 0,
        CASOrgUnit: addDatas.OrganizationUnit ? addDatas.OrganizationUnit : "",
        CASEngTypeId: addDatas.EngagemantType ? addDatas.EngagemantType : 0,
        CASEngSubTypeId: addDatas.EngagementSubType
          ? addDatas.EngagementSubType
          : 0,
        CASUserId: peopleId ? { results: peopleId } : { results: [] },
        CASStatusId: addDatas.StatusType ? addDatas.StatusType : 0,
        CASEngScopeId: addDatas.EngagementScope ? addDatas.EngagementScope : 0,
        CASCCI: addDatas.CrossCharge ? addDatas.CrossCharge : "",
        CASPrjStartDate: addDatas.ProjectStartDate
          ? addDatas.ProjectStartDate
          : null,
        CASPrjEndDate: addDatas.ProjectCompletionDate
          ? addDatas.ProjectCompletionDate
          : null,
        CASEngNotes: addDatas.EngagementNotes ? addDatas.EngagementNotes : "",
      })
      .then(() => {
        alertify.success("Record updated successfully");
        setAddDatas({
          ProjectName: "",
          Priority: null,
          CountryIBVT: null,
          OrganizationUnit: "",
          EngagemantType: null,
          EngagementSubType: 0,
          Requestor: [],
          StatusType: null,
          IDNumber: null,
          CreationDate: null,
          LastModifyDate: null,
          EngagementScope: null,
          Actions: "",
          CrossCharge: "",
          ProjectStartDate: null,
          ProjectCompletionDate: null,
          EngagementNotes: "",
        });
        setIsEdit(false);
        setReRenderTable(true);
        props.Panel(false);
      })
      .catch((error) => {
        console.log(error);
      });
  };

  // close Panel function
  const closePanel = () => {
    props.Panel(false);
    RequestorIdArr = [];
  };

  //   Tab Panel Section
  return (
    <div className={classes.PanelSection}>
      <div className={classes.PanelOverlay}>
        <div className={classes.Panel}>
          <div className={classes.topBar}>
            <button className={classes.closeBtn} onClick={closePanel}>
              <Close />
            </button>
          </div>
          <div className={classes.IconsSection}>
            <img
              className={classes.projIcon}
              src={`${ProjectIcon}`}
              width={18}
              height={18}
            />
            <img
              className={classes.lAIcon}
              src={`${LatestActionIcon}`}
              width={18}
              height={18}
            />
            {props.Admin && (
              <img
                className={classes.hSIcon}
                src={`${HoursSpentIcon}`}
                width={18}
                height={18}
              />
            )}
          </div>
          {/* Pivot - Section */}
          <Pivot
            aria-label="Basic Pivot Example"
            selectedKey={String(selectedKey)}
            onLinkClick={(e) => {
              setSelectedKey(+e.props.itemKey);
            }}
          >
            {/* Primary Section */}
            <PivotItem
              itemKey="1"
              itemIcon="SingleColumnEdit"
              headerText="PROJECT INFORMATION"
              headerButtonProps={{
                "data-order": 1,
                "data-title": "My Files Title",
              }}
            >
              <div className={classes.titleOne}>Primary Information</div>

              {/* ProjectName */}
              <div style={{ height: "125px" }} className={classes.panelInput}>
                <InputLabel required className={classes.inpLabel}>
                  Project Name:
                </InputLabel>
                <TextField
                  // inputRef={ProjectName}
                  disabled={props.Admin ? false : true}
                  value={addDatas.ProjectName}
                  className={classes.inpL}
                  id="standard-basic"
                  placeholder={`Insert here name project`}
                  variant="outlined"
                  InputLabelProps={{ shrink: false }}
                  onChange={(e) => {
                    addDatas.ProjectName = e.target.value;
                    setBtnDisabled(addDatas.ProjectName ? false : true);
                    setAddDatas({ ...addDatas });
                  }}
                />
              </div>

              <div className={`${classes.flex} ${classes.panelInput}`}>
                {/* Priority */}
                <div>
                  <InputLabel className={classes.inpLabel}>
                    Priority:
                  </InputLabel>
                  <Select
                    // inputRef={Priority}
                    disabled={props.Admin ? false : true}
                    className={classes.selectL}
                    labelId="demo-simple-select-label"
                    id="demo-simple-select"
                    value={
                      addDatas.Priority == null ? "null" : addDatas.Priority
                    }
                    onChange={(e) => {
                      addDatas.Priority = e.target.value;
                      setAddDatas({ ...addDatas });
                    }}
                    variant="outlined"
                  >
                    <MenuItem value={"null"}>Select Project Priority</MenuItem>
                    {priorityChoice.map((data) => {
                      return (
                        <MenuItem value={data.key}>
                          <div style={{ display: "flex" }}>
                            <img
                              src={`${JSON.parse(data.icon).serverRelativeUrl}`}
                              style={{ marginRight: "0.5rem" }}
                              width={24}
                              height={24}
                            />
                            <div className={classes.priorityText}>
                              {data.text}
                            </div>
                          </div>
                        </MenuItem>
                      );
                    })}
                  </Select>
                </div>

                {/* Country */}
                <div>
                  <InputLabel className={classes.inpLabel}>
                    Country/IBVT:
                  </InputLabel>
                  <Select
                    disabled={props.Admin ? false : true}
                    className={classes.selectL}
                    labelId="demo-simple-select-label"
                    id="demo-simple-select"
                    value={
                      addDatas.CountryIBVT == null
                        ? "null"
                        : addDatas.CountryIBVT
                    }
                    onChange={(e) => {
                      addDatas.CountryIBVT = e.target.value;
                      setAddDatas({ ...addDatas });
                    }}
                    variant="outlined"
                    labelWidth={0}
                  >
                    <MenuItem value={"null"}>Select Country</MenuItem>
                    {countryChoice.map((data) => {
                      return <MenuItem value={data.key}>{data.text}</MenuItem>;
                    })}
                  </Select>
                </div>
              </div>

              <div className={`${classes.flex} ${classes.panelInput}`}>
                {/* org unit */}
                <div style={{ width: "50%" }}>
                  <InputLabel className={classes.inpLabel}>
                    Organization Unit:
                  </InputLabel>
                  <TextField
                    className="orgUnitTextField"
                    disabled={props.Admin ? false : true}
                    style={{ width: "200%" }}
                    id="standard-basic"
                    variant="outlined"
                    placeholder={`Insert Organization Unit`}
                    InputLabelProps={{ shrink: false }}
                    value={addDatas.OrganizationUnit}
                    onChange={(e) => {
                      addDatas.OrganizationUnit = e.target.value;
                      setAddDatas({ ...addDatas });
                    }}
                  />
                </div>

                {/* eng type */}
                <div>
                  <InputLabel className={classes.inpLabel}>
                    Engagement Type:
                  </InputLabel>
                  <Select
                    disabled={props.Admin ? false : true}
                    className={classes.selectL}
                    labelId="demo-simple-select-label"
                    id="demo-simple-select"
                    value={
                      !addDatas.EngagemantType
                        ? "null"
                        : addDatas.EngagemantType
                    }
                    onChange={(e) => {
                      addDatas.EngagemantType = e.target.value;
                      setAddDatas({ ...addDatas });
                      setEngSubTypeChoice(
                        EngagementTypeSubArr.filter(
                          (choice) => choice.type == addDatas.EngagemantType
                        )
                      );
                    }}
                    variant="outlined"
                    labelWidth={0}
                  >
                    <MenuItem value={"null"}>Select Engagement Type</MenuItem>
                    {engTypeChoice.map((data) => {
                      return <MenuItem value={data.key}>{data.text}</MenuItem>;
                    })}
                  </Select>
                </div>
              </div>

              <div className={`${classes.flex} ${classes.panelInput}`}>
                {/* org unit */}

                {/* eng type */}
                <div style={{ width: "48%" }}>
                  <InputLabel className={classes.inpLabel}>
                    Engagement Subtype:
                  </InputLabel>
                  <Select
                    disabled={props.Admin ? false : true}
                    className={classes.selectL}
                    labelId="demo-simple-select-label"
                    id="demo-simple-select"
                    value={addDatas.EngagementSubType}
                    onChange={(e) => {
                      addDatas.EngagementSubType = e.target.value;
                      setAddDatas({ ...addDatas });
                    }}
                    variant="outlined"
                    labelWidth={0}
                  >
                    <MenuItem value={"null"}>Select Subtype</MenuItem>
                    {engSubTypeChoice.map((data) => (
                      <MenuItem value={data.key}>{data.text}</MenuItem>
                    ))}
                  </Select>
                </div>
              </div>
              {/* Requestor */}
              <div className={classes.panelInput}>
                <InputLabel className={classes.inpLabel}>Requestor:</InputLabel>
                <div className={classes.panelPPicker}>
                  <PeoplePicker
                    disabled={props.Admin ? false : true}
                    context={props.context}
                    placeholder={`Insert Requestor people`}
                    personSelectionLimit={10}
                    showtooltip={true}
                    ensureUser={true}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    onChange={(e) => {
                      RequestorArr = [];
                      RequestorIdArr = [];
                      if (e.length > 0) {
                        e.forEach((data) => {
                          RequestorArr.push(data.secondaryText);
                          RequestorIdArr.push(data.id);
                        });
                        setPeopleId(RequestorIdArr);
                      } else {
                        setPeopleId(RequestorIdArr);
                      }
                      addDatas.Requestor = e.length ? RequestorArr : [];
                      setAddDatas({ ...addDatas });
                    }}
                    defaultSelectedUsers={addDatas.Requestor}
                    required={true}
                  />
                </div>
              </div>

              <div
                className={`${classes.flex} ${classes.panelInput}`}
                style={{ marginTop: "1.5rem" }}
              >
                {/* status type */}
                <div>
                  <InputLabel className={classes.inpLabel}>
                    Status Type:
                  </InputLabel>
                  <Select
                    disabled={props.Admin ? false : true}
                    className={classes.selectL}
                    labelId="demo-simple-select-label"
                    id="demo-simple-select"
                    value={!addDatas.StatusType ? "null" : addDatas.StatusType}
                    onChange={(e) => {
                      addDatas.StatusType = e.target.value;
                      setAddDatas({ ...addDatas });
                    }}
                    variant="outlined"
                    labelWidth={0}
                  >
                    <MenuItem value={"null"}>Select Status Type</MenuItem>
                    {statusTypeChoice.map((data) => {
                      return (
                        <MenuItem value={data.key}>
                          <div
                            className={classes.listItem}
                            style={{
                              color:
                                data.text.toLowerCase() == "in progress"
                                  ? "#359942"
                                  : data.text.toLowerCase() ==
                                    "waiting for feedback"
                                  ? "#f5944e"
                                  : data.text.toLowerCase() == "lead"
                                  ? "#f24998"
                                  : data.text.toLowerCase() == "parked"
                                  ? "#999999"
                                  : data.text.toLowerCase() == "closed"
                                  ? "#1c75bc"
                                  : data.text.toLowerCase() == "canceled"
                                  ? "#7e2e7a"
                                  : "#000",
                            }}
                          >
                            <div
                              style={{
                                background:
                                  data.text.toLowerCase() == "in progress"
                                    ? "#359942"
                                    : data.text.toLowerCase() ==
                                      "waiting for feedback"
                                    ? "#f5944e"
                                    : data.text.toLowerCase() == "lead"
                                    ? "#f24998"
                                    : data.text.toLowerCase() == "parked"
                                    ? "#999999"
                                    : data.text.toLowerCase() == "closed"
                                    ? "#1c75bc"
                                    : data.text.toLowerCase() == "canceled"
                                    ? "#7e2e7a"
                                    : "#000",
                              }}
                            ></div>
                            {data.text}
                          </div>
                        </MenuItem>
                      );
                    })}
                  </Select>
                </div>

                {/* id num */}
                {isEdit && (
                  <div>
                    <InputLabel className={classes.inpLabel}>
                      ID Number:
                    </InputLabel>
                    <div className="IdDropdown">
                      {/* <Select
                        disabled={true}
                        className={classes.selectL}
                        labelId="demo-simple-select-label"
                        id="demo-simple-select"
                        variant="outlined"
                        labelWidth={0}
                        value={props.Edit.item}
                      >
                        <MenuItem value={props.Edit.item}>
                          {props.Edit.itemta}
                        </MenuItem>
                      </Select> */}
                      <TextField
                        disabled={true}
                        value={
                          addDatas.IDNumber != 0 ? addDatas.IDNumber : null
                        }
                        style={{ width: "200%" }}
                        id="standard-basic"
                        variant="outlined"
                      />
                    </div>
                  </div>
                )}

                {/* creation date */}
                {isEdit && (
                  <div>
                    <InputLabel className={classes.inpLabel}>
                      Creation Date:
                    </InputLabel>
                    <DatePicker
                      disabled={true}
                      className={classes.dateL}
                      formatDate={(date: Date): string => {
                        let arrDate = date.toLocaleDateString().split("/");
                        let selectedDate = `${
                          +arrDate[1] < 10 ? "0" + arrDate[1] : arrDate[1]
                        }/${
                          +arrDate[0] < 10 ? "0" + arrDate[0] : arrDate[0]
                        }/${arrDate[2].toString().substr(-2)}`;
                        return selectedDate;
                        // return (
                        //   date.getDate() +
                        //   "/" +
                        //   (date.getMonth() + 1) +
                        //   "/" +
                        //   date.getFullYear()
                        // );
                      }}
                      value={
                        addDatas.CreationDate
                          ? new Date(addDatas.CreationDate)
                          : null
                      }
                      onSelectDate={(e) => {
                        console.log(e.toISOString());
                        // handleValue("CreationDate", e.toISOString());
                      }}
                    />
                  </div>
                )}

                {/* last mod date */}
                {isEdit && (
                  <div>
                    <InputLabel className={classes.inpLabel}>
                      Last Modify Date:
                    </InputLabel>
                    <DatePicker
                      disabled={true}
                      className={classes.dateL}
                      formatDate={(date: Date): string => {
                        let arrDate = date.toLocaleDateString().split("/");
                        let selectedDate = `${
                          +arrDate[1] < 10 ? "0" + arrDate[1] : arrDate[1]
                        }/${
                          +arrDate[0] < 10 ? "0" + arrDate[0] : arrDate[0]
                        }/${arrDate[2].toString().substr(-2)}`;
                        return selectedDate;
                        // return (
                        //   date.getDate() +
                        //   "/" +
                        //   (date.getMonth() + 1) +
                        //   "/" +
                        //   date.getFullYear()
                        // );
                      }}
                      value={
                        addDatas.LastModifyDate
                          ? new Date(addDatas.LastModifyDate)
                          : null
                      }
                      onSelectDate={(e) => {
                        console.log(e.toISOString());
                        // handleValue("LastModifyDate", e.toISOString());
                      }}
                    />
                  </div>
                )}
              </div>

              <div className={classes.titleOne}>Details Information</div>

              <div className={`${classes.flex} ${classes.panelInput}`}>
                {/* Eng scope */}
                <div>
                  <InputLabel className={classes.inpLabel}>
                    Engagement Scope:
                  </InputLabel>
                  <Select
                    disabled={props.Admin ? false : true}
                    className={classes.selectL}
                    labelId="demo-simple-select-label"
                    id="demo-simple-select"
                    value={
                      !addDatas.EngagementScope
                        ? "null"
                        : addDatas.EngagementScope
                    }
                    onChange={(e) => {
                      addDatas.EngagementScope = e.target.value;
                      setAddDatas({ ...addDatas });
                    }}
                    variant="outlined"
                    labelWidth={0}
                  >
                    <MenuItem value={"null"}>Select Engagement Scope</MenuItem>
                    {engScopeChoice.map((data) => {
                      return <MenuItem value={data.key}>{data.text}</MenuItem>;
                    })}
                  </Select>
                </div>
              </div>

              {/* Cross charge information */}
              <div style={{ height: "130px" }} className={classes.panelInput}>
                <InputLabel className={classes.inpLabel}>
                  Cross charge information:
                </InputLabel>
                <TextField
                  className={classes.inpL}
                  id="standard-basic"
                  variant="outlined"
                  placeholder={`Insert Charge Information`}
                  InputLabelProps={{ shrink: false }}
                  value={addDatas.CrossCharge}
                  onChange={(e) => {
                    addDatas.CrossCharge = e.target.value;
                    setAddDatas({ ...addDatas });
                  }}
                />
              </div>

              <div
                style={{ height: "110px" }}
                className={`${classes.flex} ${classes.panelInput}`}
              >
                {/* project start date */}
                <div>
                  <InputLabel className={classes.inpLabel}>
                    Project Start Date:
                  </InputLabel>
                  <DatePicker
                    disabled={props.Admin ? false : true}
                    className={classes.dateL}
                    placeholder={`Insert Date`}
                    formatDate={(date: Date): string => {
                      let arrDate = date.toLocaleDateString().split("/");
                      let selectedDate = `${
                        +arrDate[1] < 10 ? "0" + arrDate[1] : arrDate[1]
                      }/${
                        +arrDate[0] < 10 ? "0" + arrDate[0] : arrDate[0]
                      }/${arrDate[2].toString().substr(-2)}`;
                      return selectedDate;
                      // return (
                      //   date.getDate() +
                      //   "/" +
                      //   (date.getMonth() + 1) +
                      //   "/" +
                      //   date.getFullYear()
                      // );
                    }}
                    value={
                      addDatas.ProjectStartDate
                        ? new Date(addDatas.ProjectStartDate)
                        : null
                    }
                    onSelectDate={(e) => {
                      addDatas.ProjectStartDate = e.toISOString();
                      setAddDatas({ ...addDatas });
                    }}
                  />
                </div>

                {/* project completion date */}
                <div>
                  <InputLabel className={classes.inpLabel}>
                    Project Completion Date:
                  </InputLabel>
                  <DatePicker
                    disabled={props.Admin ? false : true}
                    className={classes.dateL}
                    placeholder={`Insert Date`}
                    formatDate={(date: Date): string => {
                      let arrDate = date.toLocaleDateString().split("/");
                      let selectedDate = `${
                        +arrDate[1] < 10 ? "0" + arrDate[1] : arrDate[1]
                      }/${
                        +arrDate[0] < 10 ? "0" + arrDate[0] : arrDate[0]
                      }/${arrDate[2].toString().substr(-2)}`;
                      return selectedDate;
                      // return (
                      //   date.getDate() +
                      //   "/" +
                      //   (date.getMonth() + 1) +
                      //   "/" +
                      //   date.getFullYear()
                      // );
                    }}
                    value={
                      addDatas.ProjectCompletionDate
                        ? new Date(addDatas.ProjectCompletionDate)
                        : null
                    }
                    onSelectDate={(e) => {
                      addDatas.ProjectCompletionDate = e.toISOString();
                      setAddDatas({ ...addDatas });
                    }}
                  />
                </div>
              </div>

              {/* Engagement Notes */}
              <div style={{ height: "130px" }} className={classes.panelInput}>
                <InputLabel className={classes.inpLabel}>
                  Engagement Notes:
                </InputLabel>
                <TextField
                  disabled={props.Admin ? false : true}
                  className={classes.inpL}
                  id="standard-basic"
                  variant="outlined"
                  placeholder={`Insert Engagement Notes`}
                  InputLabelProps={{ shrink: false }}
                  value={addDatas.EngagementNotes}
                  onChange={(e) => {
                    addDatas.EngagementNotes = e.target.value;
                    setAddDatas({ ...addDatas });
                  }}
                />
              </div>

              {/* Button Section */}
              <div className={classes.actions}>
                <button
                  className={`${classes.publishBtn} ${
                    addDatas.ProjectName == "" ? classes.publishBtnDisabled : ""
                  }`}
                  disabled={addDatas.ProjectName == "" ? true : false}
                  onClick={() => {
                    isEdit ? UpdateListItem() : AddListItems();
                  }}
                >
                  PUBLISH
                </button>
              </div>
            </PivotItem>

            {/* Latest Action Section */}
            <PivotItem headerText="LATEST ACTION" itemIcon="Chat" itemKey="2">
              <LatestAction
                renderProject={props.renderProject}
                Edit={props.Edit}
                Latest={latestSec}
                sp={props.sp}
                Admin={props.Admin}
                editLatest={props.latest}
              />
            </PivotItem>

            {/* Hours spent Section */}
            {props.Admin ? (
              <PivotItem
                headerText="HOURS SPENT"
                itemIcon="DateTime"
                itemKey="3"
              >
                <HoursSpent
                  renderProject={props.renderProject}
                  Latest={latestSec}
                  sp={props.sp}
                  Edit={props.Edit}
                />
              </PivotItem>
            ) : (
              ""
            )}
          </Pivot>
          {/* Pivot - Section */}
        </div>
      </div>
    </div>
  );
};
export default Panel;
