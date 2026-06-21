Redmine::Plugin.register :employee_workload_report do
  name 'Employee Workload Report'
  author 'Student'
  description 'Отчет по выработке сотрудников'
  version '0.0.1'

  menu :top_menu,
       :employee_workload_report,
       { controller: 'employee_workload_report', action: 'index' },
       caption: 'Выработка'
end
Rails.application.config.i18n.load_path +=
  Dir[File.expand_path('config/locales/*.yml', __dir__)]
