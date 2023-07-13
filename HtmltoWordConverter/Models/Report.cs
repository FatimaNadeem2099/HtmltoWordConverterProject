using Microsoft.AspNetCore.Mvc.Rendering;
using System.Collections.Generic;
using System;

namespace HtmltoWordConverter.Models
{
    public class Report
    {

        public int clientId { get; set; }
        public int companyId { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        public List<DepartmentVM> DepartmentSection { get; set; }
        public List<ObservationVM> ObservationVM { get; set; }
        public List<products> products { get; set; }
        public List<string> language { get; set; }
        public Boolean mergeObservations { get; set; }
    }

    public class DepartmentVM
    {
        public int departmentId { get; set; }
        public int SubDepartmentId { get; set; }

    }
    public class ObservationVM
    {
        public int ObservationId { get; set; }
        public string ObservationRisk { get; set; }
        public List<string> ObservationSamples { get; set; }

    }
    public class Client
    {
        public int clientId { get; set; }
        public string Name { get; set; }
    }

    public class Observation1
    {
        public int observationId { get; set; }
        public string Name { get; set; }
    }
    public class Company
    {
        public int companyId { get; set; }
        public string Name { get; set; }
    }
    public class Department
    {
        public int departmentId { get; set; }
        public string Name { get; set; }
    }
    public class SubDepartment
    {
        public int SubDepartmentId { get; set; }
        public int departmentId { get; set; }
        public string Name { get; set; }
    }

    public class products
    {
        public int year { get; set; }
        public int population { get; set; }
        public int sample { get; set; }
    }

    public class CascadingDropdownViewModel
    {
        public int SelecteddepartmentId { get; set; }
        public int SelectedSubDepartmentId { get; set; }
        public IEnumerable<SelectListItem> Departments { get; set; }
    }



    public class CascadingDropdownsViewModel1
    {
        public int SelectedCountryId { get; set; }
        public int SelectedCityId { get; set; }
        public IEnumerable<SelectListItem> Countries { get; set; }
    }

    public class DynamicFormElement
    {
        public int SelectedCountryId { get; set; }
        public int SelectedCityId { get; set; }
    }
}

