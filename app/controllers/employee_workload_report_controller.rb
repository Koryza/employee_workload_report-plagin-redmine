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
end
