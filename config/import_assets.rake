require 'rubyXL'
require 'pp'
require 'pry'

namespace :dynamo do
  desc 'Load non-standard WU Asset data from VCM into the database from MS Excel: rake dynamo:load_vmc_data[file_path,customer_id]'
  task :import_assets, [:file_path, :customer_id, :user_id] => :environment do |task, args|
    raise 'Need file path as first argument [file_path,customer_id, user_id]'    if args.file_path.blank?
    raise 'Need customer_id as second argument [file_path,customer_id, user_id]' if args.customer_id.blank?
    raise 'Need user as third argument [file_path,customer_id, user_id]'         if args.user_id.blank?
    raise 'File should exist so we can do things with it!'                       unless File.exists? args.file_path

    def check_asset_type(row, existing_asset=nil)
      if existing_asset.nil?
        return VirtualServer if is_virtual_server? row
        return Server
      else
        return existing_asset.type
      end
    end

    def check_asset_exists(row)
      asset = Asset.where(
                       name: clean_hostname(row),
                       customer: @customer
      ).first
      return asset
    end

    def check_asset_valid(asset)
      if asset.valid? == false && asset.errors[:serial].count > 0
        puts "Nil'd Serial for Asset: '#{ asset.name }' which had serial: '#{ asset.serial }'"
        asset.serial = nil
      end
      return asset
    end

    def check_asset_updated(asset, row)
      #binding.pry if asset.name == 'DEVWWW213'      # DEBUG
      asset = check_value_operating_system(asset, row)
      return asset
    end

    def check_value_disposition(asset, row)
      return asset if (row[12].nil? || row[12] == '')
      @asset_disposition_ids ||= AssetDisposition.select(:name).all.map(&:name)
      disposition = row[12].rstrip

      unless @asset_disposition_ids.include?(disposition)
        raise "Asset Disposition: '#{ disposition }' does not exist or is otherwise invalid!"
      end

      #puts "asset-disposition-#{ disposition }".parameterize.to_sym # DEBUG
      disposition_id = Rails.cache.fetch("asset-disposition-#{ disposition }".parameterize.to_sym) {
        AssetDisposition.where(name: disposition).first!
      }

      check_asset_valid(asset)
      if ((asset.asset_disposition_id.nil?) || (asset.asset_disposition_id != disposition_id))  # if the Asset has no Make/Model but the Spreadsheet does...
        asset.update_attributes!(asset_disposition_id: disposition_id)
      elsif (asset.asset_disposition_id != disposition_id)                             # if the Asset has an Make/Model but is different than the Spreadsheet's M&M...
        asset.update_attributes!(asset_disposition_id: disposition_id)
      end
      return asset
    end

    def check_value_location(asset, row)
      begin
        loc = find_or_create_location(row)
        return asset if loc.nil?

        check_asset_valid(asset)
        if (asset.location_id.nil?) && (loc.nil? == false)   # if the Asset has no Location but the Spreadsheet does...
          #puts "Updated LOC for #{ asset.name } to #{ loc.name }"     # DEBUG
          asset.update_attributes!(location: loc)
        elsif (asset.location != loc)                        # if the Asset has an Location but is different than the Spreadsheet's Location...
          #puts "Updated LOC for #{ asset.name } to #{ loc.name }"     # DEBUG
          asset.update_attributes!(location: loc)
        end
        return asset
      rescue Exception => e
        #binding.pry   # DEBUG
        raise e
      end
    end

    def check_value_make_model(asset, row)
      return asset if (row[6].nil? || row[6] == '') || (row[7].nil? || row[7] == '')
      manufacturer = row[6].upcase.rstrip
      model        = row[7].upcase.rstrip

      check_asset_valid(asset)
      if ((asset.ff_vendor.nil?) || (asset.ff_vendor != manufacturer)) || ((asset.ff_model.nil?) || (asset.ff_model != model))  # if the Asset has no Make/Model but the Spreadsheet does...
        asset.update_attributes!(ff_vendor: manufacturer, ff_model: model)
      elsif (asset.ff_vendor != manufacturer) || (asset.ff_model != model)                                                      # if the Asset has an Make/Model but is different than the Spreadsheet's M&M...
        asset.update_attributes!(ff_vendor: manufacturer, ff_model: model)
      end
      return asset
    end

    def check_value_operating_system(asset, row)
      os = find_or_create_os(row)
      return asset if os.nil?                                     # if the Spreadsheet has no OS, do nothing to the Asset

      check_asset_valid(asset)
      if (asset.operating_system_id.nil?) && (os.nil? == false)   # if the Asset has no OS but the Spreadsheet does...
        asset.update_attributes!(operating_system: os)
      elsif (asset.operating_system != os)                        # if the Asset has an OS but is different than the Spreadsheet's OS...
        asset.update_attributes!(operating_system: os)
      end
      return asset
    end

    def check_value_serial(asset, row)
      begin
        serial = clean_serial(row)
        return asset if serial.nil?

        check_asset_valid(asset)
        if (asset.serial.nil?) && (serial.nil? == false)   # if the Asset has no VE but the Spreadsheet does...
          asset.update_attributes!(serial: serial)
        elsif (asset.serial != serial)                        # if the Asset has an VE but is different than the Spreadsheet's VE...
          asset.update_attributes!(serial: serial)
        end
        return asset
      rescue Exception => e
        #binding.pry   # DEBUG
        raise e
      end
    end

    def check_value_tracker_id(asset, row)
      begin
        tracker_id = clean_tracker_id(row)
        return asset if tracker_id.nil?

        check_asset_valid(asset)
        if (asset.asset_tracker.nil?) && (tracker_id.nil? == false)   # if the Asset has no VE but the Spreadsheet does...
          asset.update_attributes!(asset_tracker: tracker_id)
        elsif (asset.asset_tracker != tracker_id)                        # if the Asset has an VE but is different than the Spreadsheet's VE...
          asset.update_attributes!(asset_tracker: tracker_id)
        end
        return asset
      rescue Exception => e
        #binding.pry   # DEBUG
        raise e
      end
    end

    def check_value_virtual_environment(asset, row)
      begin
        virtual_environment = find_or_create_virtual_environment(row)
        return asset if virtual_environment.nil?

        check_asset_valid(asset)
        if (asset.virtual_environment_id.nil?) && (virtual_environment.nil? == false)   # if the Asset has no VE but the Spreadsheet does...
          asset.update_attributes!(virtual_environment: virtual_environment)
        elsif (asset.virtual_environment != virtual_environment)                        # if the Asset has an VE but is different than the Spreadsheet's VE...
          asset.update_attributes!(virtual_environment: virtual_environment)
        end
        return asset
      rescue Exception => e
        #binding.pry   # DEBUG
        raise e
      end
    end

    def clean_hostname(row)
      begin
        cleaned = row[0].rstrip.upcase.gsub(' ', '-').gsub('.', '-')
        if cleaned.match(/\(/)
          cleaned = cleaned.split('(').first
        end
        return cleaned
      rescue Exception => e
        #binding.pry   # DEBUG
        raise e
      end
    end

    def clean_ip_addresses(row)
      return Array.new if row[1].nil?
      ips = row[1].split(',')
      ips.each do |ip|
        ip.rstrip!
      end
      return ips
    end

    def clean_serial(row)
      return nil if row[8].nil? || row[8].rstrip == ''
      return row[8].rstrip.upcase
    end

    def create_operating_system(row)
      begin
        return os = find_or_create_os(row)
      rescue Exception => e
        #pp row      # DEBUG
        binding.pry # DEBUG
        raise e
      end
    end

    def clean_tracker_id(row)
      return row[1].rstrip.upcase
    end

    def create_vendor(name)
      if name.include?('VMware')
        polished_name = 'VMWARE'
      elsif name.upcase.match('WINDOWS')
        polished_name = 'MICROSOFT'
      elsif name.upcase.match('CISCO')
        polished_name = 'CISCO'
      elsif name.upcase.match('DELL')
        polished_name = 'DELL'
      elsif name.upcase.match('HEWLETT-PACKARD')
        polished_name = 'HP'
      else
        polished_name = name.upcase
      end
      Vendor.where(name: polished_name).first_or_create!
    end

    def find_or_create_location(row)
      begin
        return nil if row[15].nil? || row[15].rstrip == ''
        loc_field  =  row[15].rstrip.upcase

        loc = Location.where(
            name: loc_field,
            customer: @customer
        ).first_or_create!
        return loc
      rescue Exception => e
        #binding.pry # DEBUG
        raise e
      end
    end

    def find_or_create_os(row)
      return nil if row[13].nil? || row[13] == '0' || row[13] == 0
      os_field   = row[13].rstrip.upcase
      os = PEN::NormalizeOperatingSystem.normalize(name: os_field)
      return os
    end

    def find_or_create_virtual_environment(row)
      begin
        return nil if row[4].nil? || row[4].rstrip == ''
        ve_field   =  row[4].rstrip.upcase

        virtual_environment = VirtualEnvironment.where(
            name: ve_field,
            customer: @customer
        ).first_or_create!
        return virtual_environment
      rescue Exception => e
        #binding.pry # DEBUG
        raise e
      end
    end

    def is_migration_complete?(row)
      return false unless row[25]
      return true if row[25].upcase == 'YES'
      return false
    end

    def is_virtual_server?(row)
      return true if row[5].to_s.match(/^VM/)
      return false
    end

    def modify_description(row)
      return nil if row[6].nil? && row[7].nil? && row[1].nil?
      desc = "<P>Model Info: #{ row[6] } #{ row[7] }</P>" +
             "<P><EM>(Data Source: Spreadsheet provided by Emanuel R., Forsythe)</EM></P>"
      desc = "<P>#{ row[1] }</P>" + desc unless row[1].nil?
      return desc
    end

    def process_worksheet_assets(worksheet)
      puts 'Processing Assets...'
      worksheet.shift                                             # Discard the header row
      i = 0
      worksheet.each do |row|
        begin
          next if row.nil? || (row[0].nil? && row[2].nil?)
          i += 1
          puts i    # DEBUG
          existing_asset = check_asset_exists(row)
          asset_type     = check_asset_type(row, existing_asset)
          if existing_asset.nil?
            created_asset  = asset_type.where(
                name: clean_hostname(row).upcase,
                customer: @customer
            ).first_or_create!(
                description: modify_description(row),
                operating_system: create_operating_system(row),
                migration_complete: is_migration_complete?(row)
            )
            existing_asset = created_asset
            #pp created_asset  # DEBUG
          end
          existing_asset = check_asset_updated(existing_asset, row)                   # TODO this is outdated
          existing_asset = check_value_tracker_id(existing_asset, row)
          existing_asset = check_value_disposition(existing_asset, row)
          existing_asset = check_value_location(existing_asset, row)
          existing_asset = check_value_serial(existing_asset, row)
          existing_asset = check_value_virtual_environment(existing_asset, row)
          existing_asset = check_value_make_model(existing_asset, row)

            #machine_model         = row[9]
            #os                    = create_operating_system row
            #server_type_detail_id = machine_model == 'NULL' ? nil : create_server_type_detail(row).id
            #created_asset         = nil
            #asset_description_append = "[Data source: 'Dynamo Asset Template v1.0' data]"

            #puts '*' * 20                                          # DEBUG
            #puts created_asset.description                         # DEBUG
            #pp created_asset                                       # DEBUG
            #pp created_asset.operating_system                      # DEBUG
            #pp created_asset.server_type_detail                    # DEBUG
            #pp created_asset.server_type_detail.server_type        # DEBUG
        rescue Exception => e
          #binding.pry # DEBUG
          puts e.message
          puts e.backtrace
          raise e
        end
      end
    end

    def process_worksheet_ips(worksheet)
      puts 'Processing IPs...'
      worksheet.shift                                             # Discard the header row
      worksheet.each do |row|
        begin
          next if row.nil? || (row[0].nil? || row[1].nil?)         # we've actually seen this...
          asset_name    = clean_hostname(row)
          related_asset = Asset.where(
                            name: asset_name,
                            customer: @customer
          ).first
          if related_asset.nil?
            puts "Couldn't find Asset w/ name: '#{ asset_name } while processing IPs. Skipping."
            next
          end

          ips = clean_ip_addresses(row)
          ips.each do |ip_address|
            ip = IpAddress.new
            ip.customer = @customer
            ip.name     = ip_address.lstrip.rstrip
            unless ip.valid?
              unless (ip.errors.count == ip.errors[:name].count) && (ip.errors.messages[:name].include?('has already been taken'))
                # IpAddress#name is expected to not validate as unique due to how we're doing this. This specific validation failure can be ignored.
                puts "WARN: IP '#{ ip.name }' did not validate for asset '#{ related_asset.name }'!"
                pp ip.errors # DEBUG
                next
              end
            end
            # the above is just to weed out any data that doesn't confirm to IpAddress validations.

            ip = IpAddress.where(name: ip.name, customer: ip.customer).first_or_create!
            related_asset.ip_addresses.push ip unless related_asset.ip_addresses.include?(ip)
          end
        rescue Exception => e
          binding.pry # DEBUG
          puts e.message
          puts e.backtrace
          raise e
        end
      end
    end

    def upcase_existing_assets
      puts 'Upcasing Existing Asset Names not already upper-cased!'     # DEBUG
      Asset.where(customer: @customer).find_each do |asset|
        upcased_name = asset.name.upcase
        unless asset.name == upcased_name
          asset.name = upcased_name
          asset.save!
          puts "Upcased Asset ID '#{ asset.id }' to '#{ asset.name }'."
        end
      end
    end

    ### Main Program ###

    @customer          = Customer.find(args.customer_id.to_i)
    puts "Initial Asset count: #{ Asset.where(customer: @customer).count }"            # DEBUG
    puts "Initial LOC count: #{ Location.where(customer: @customer).count }"           # DEBUG
    puts "Initial VE count: #{ VirtualEnvironment.where(customer: @customer).count }"  # DEBUG
    #puts "Asset 'BACSANDBOX01' has VE ID: '#{ @customer.assets.where(name: 'BACSANDBOX01').first!.virtual_environment_id }'"  # DEBUG
    #puts "Asset 'BACSANDBOX01' has LOC ID: '#{ @customer.assets.where(name: 'BACSANDBOX01').first!.location_id }'"            # DEBUG
    upcase_existing_assets
    user              = User.find(args.user_id.to_i)
    workbook          = RubyXL::Parser.parse(args.file_path)
    worksheet_assets  = workbook['assets']
    worksheet_assets  = worksheet_assets.extract_data
    #pp Asset.where(name: 'DEVMTXWWW01', customer: @customer).first!   # DEBUG
    #pp Asset.where(name: 'DEVWWW213', customer: @customer).first!     # DEBUG
    #pp Asset.where(name: 'N0ASREC001', customer: @customer).first!    # DEBUG
    puts '*' * 20                               # DEBUG
    process_worksheet_assets worksheet_assets
    #pp Asset.where(name: 'DEVMTXWWW01', customer: @customer).first!   # DEBUG
    #pp Asset.where(name: 'DEVWWW213', customer: @customer).first!     # DEBUG
    #pp Asset.where(name: 'N0ASREC001', customer: @customer).first!    # DEBUG
    worksheet_assets  = nil         # we're done w/ this

    worksheet_ips     = workbook['asset_ips']
    worksheet_ips     = worksheet_ips.extract_data
    process_worksheet_ips worksheet_ips

    puts "Final Asset count: #{ Asset.where(customer: @customer).count }"               # DEBUG
    puts "\tServer count: #{ Server.where(customer: @customer).count }"                 # DEBUG
    puts "\tVirtual Server count: #{ VirtualServer.where(customer: @customer).count }"  # DEBUG
    puts "Final LOC count: #{ Location.where(customer: @customer).count }"              # DEBUG
    puts "Final VE count: #{ VirtualEnvironment.where(customer: @customer).count }"     # DEBUG
    #puts "Asset 'BACSANDBOX01' has VE ID: '#{ @customer.assets.where(name: 'BACSANDBOX01').first!.virtual_environment_id }'"  # DEBUG
    #puts "Asset 'BACSANDBOX01' has LOC ID: '#{ @customer.assets.where(name: 'BACSANDBOX01').first!.location_id }'"            # DEBUG
  end
end

