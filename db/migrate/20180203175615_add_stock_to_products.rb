class AddStockToProducts < ActiveRecord::Migration[5.0]
  def change
    add_column :products, :stock_qty, :integer
  end
end