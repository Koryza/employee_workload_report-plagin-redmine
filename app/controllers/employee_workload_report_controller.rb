class EmployeeWorkloadReportController < ApplicationController
  def index
    @from = params[:from]
    @to = params[:to]
    @project_id = params[:project_id]
    @user_id = params[:user_id]

    @projects = Project.order(:name)
    @users = User.active.order(:lastname)

    entries = TimeEntry.joins(:user, :project)

    entries = entries.where("spent_on >= ?", @from) if @from.present?
    entries = entries.where("spent_on <= ?", @to) if @to.present?
    entries = entries.where(project_id: @project_id) if @project_id.present?
    entries = entries.where(user_id: @user_id) if @user_id.present?

    @total_hours = entries.sum(:hours)

    @employee_count = entries.select(:user_id).distinct.count
    @project_count = entries.select(:project_id).distinct.count

    @user_totals = entries.group(:user_id).sum(:hours)

    @data = entries
      .select("
        users.id as user_id,
        users.firstname,
        users.lastname,
        projects.name AS project_name,
        SUM(time_entries.hours) AS total_hours
      ")
      .group("users.id, projects.id")
      .order("users.lastname")
  end

  def export
    require 'caxlsx'

    from = params[:from]
    to = params[:to]
    project_id = params[:project_id]
    user_id = params[:user_id]

    entries = TimeEntry.joins(:user, :project)

    entries = entries.where("spent_on >= ?", from) if from.present?
    entries = entries.where("spent_on <= ?", to) if to.present?
    entries = entries.where(project_id: project_id) if project_id.present?
    entries = entries.where(user_id: user_id) if user_id.present?

    user_totals = entries.group(:user_id).sum(:hours)

    data = entries
      .select("
        users.id as user_id,
        users.firstname,
        users.lastname,
        projects.name AS project_name,
        SUM(time_entries.hours) AS total_hours
      ")
      .group("users.id, projects.id")
      .order("users.lastname")

    package = Axlsx::Package.new

    package.workbook.add_worksheet(name: "Отчет") do |sheet|
      sheet.add_row ["Сотрудник", "Проект", "Часы", "% участия"]

      data.each do |row|
        total_user = user_totals[row.user_id] || 0

        percent =
          if total_user > 0
            ((row.total_hours.to_f / total_user) * 100).round(2)
          else
            0
          end

        sheet.add_row [
          "#{row.firstname} #{row.lastname}",
          row.project_name,
          row.total_hours,
          "#{percent}%"
        ]
      end
    end

    send_data(
      package.to_stream.read,
      filename: "employee_workload_report.xlsx",
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
  end
  def export_year
    require 'caxlsx'

    year = (params[:year] || Date.today.year).to_i
    package = Axlsx::Package.new

    package.workbook.add_worksheet(name: "Сводка") do |sheet|
      total_hours_year = TimeEntry.where(spent_on: Date.new(year, 1, 1)..Date.new(year, 12, 31)).sum(:hours)

      employee_count = TimeEntry.where(spent_on: Date.new(year, 1, 1)..Date.new(year, 12, 31))
                                .select(:user_id)
                                .distinct
                                .count

      project_count = TimeEntry.where(spent_on: Date.new(year, 1, 1)..Date.new(year, 12, 31))
                               .select(:project_id)
                               .distinct
                               .count

      average_hours = employee_count > 0 ? (total_hours_year.to_f / employee_count).round(2) : 0

      top_user = TimeEntry.joins(:user)
                          .where(spent_on: Date.new(year, 1, 1)..Date.new(year, 12, 31))
                          .group("users.firstname, users.lastname")
                          .sum(:hours)
                          .max_by { |_, hours| hours }

      top_project = TimeEntry.joins(:project)
                             .where(spent_on: Date.new(year, 1, 1)..Date.new(year, 12, 31))
                             .group("projects.name")
                             .sum(:hours)
                             .max_by { |_, hours| hours }

      average_load = ((average_hours / 2016.0) * 100).round(2)

      sheet.add_row ["Годовой отчет по выработке сотрудников"]
      sheet.add_row []
      sheet.add_row ["Год", year]
      sheet.add_row ["Всего сотрудников", employee_count]
      sheet.add_row ["Всего проектов", project_count]
      sheet.add_row ["Всего часов за год", total_hours_year]
      sheet.add_row ["Среднее часов на сотрудника", average_hours]

      if top_user
        sheet.add_row [
          "Самый активный сотрудник",
          "#{top_user[0][0]} #{top_user[0][1]} (#{top_user[1]} ч.)"
        ]
      end

      if top_project
        sheet.add_row [
          "Самый активный проект",
          "#{top_project[0]} (#{top_project[1]} ч.)"
        ]
      end

      sheet.add_row ["Средняя загрузка сотрудников", "#{average_load}%"]
    end

    months = [
      "Январь", "Февраль", "Март", "Апрель",
      "Май", "Июнь", "Июль", "Август",
      "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    ]

    months.each_with_index do |month_name, index|
      package.workbook.add_worksheet(name: month_name) do |sheet|
        sheet.add_row ["Отчет за #{month_name} #{year}"]
        sheet.add_row []
        sheet.add_row ["Сотрудник", "Проект", "Часы", "% участия", "План часов", "% выполнения плана"]

        month_start = Date.new(year, index + 1, 1)
        month_end = month_start.end_of_month

        entries = TimeEntry.joins(:user, :project)
                           .where(spent_on: month_start..month_end)

        user_totals = entries.group(:user_id).sum(:hours)

        data = entries.select("users.id as user_id, users.firstname, users.lastname, projects.name AS project_name, SUM(time_entries.hours) AS total_hours")
                      .group("users.id, projects.id")
                      .order("users.lastname")

        plan_hours = 168

        data.each do |row|
          total_user = user_totals[row.user_id] || 0

          participation = total_user > 0 ? ((row.total_hours.to_f / total_user) * 100).round(2) : 0
          plan_percent = ((row.total_hours.to_f / plan_hours) * 100).round(2)

          sheet.add_row [
            "#{row.firstname} #{row.lastname}",
            row.project_name,
            row.total_hours,
            "#{participation}%",
            plan_hours,
            "#{plan_percent}%"
          ]
        end
      end
    end

    package.workbook.add_worksheet(name: "Проекты по месяцам") do |sheet|
      sheet.add_row(["Проект"] + months)

      Project.order(:name).each do |project|
        row = [project.name]

        (1..12).each do |month|
          month_start = Date.new(year, month, 1)
          month_end = month_start.end_of_month

          hours = TimeEntry.where(project_id: project.id)
                           .where(spent_on: month_start..month_end)
                           .sum(:hours)

          row << hours
        end

        sheet.add_row row
      end
    end

    send_data(
      package.to_stream.read,
      filename: "employee_workload_#{year}.xlsx",
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
  end
end
