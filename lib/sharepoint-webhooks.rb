# frozen_string_literal: true
module Sharepoint
  class Subscription < Sharepoint::Object
    belongs_to :list

    field 'clientState', default: ''
    field 'expirationDateTime'
    field 'notificationUrl'
    field 'resource'

    def initialize(site, data = {})
      super site, data
    end

    def create
      @site.query :post, create_uri, @data.to_json, true do |curl|
        curl.headers['Accept'] = 'application/json;odata=nometadata'
        curl.headers['Content-Type'] = curl.headers['Accept']
      end
    end

    def destroy
      @site.query :post, resource_uri do |curl|
        curl.headers['X-HTTP-Method'] = 'DELETE'
      end
    end

    private

    def resource_uri
      __metadata['uri'][/^.*_api\//i] +
      "web/lists(guid'#{@data['resource']}')/" \
      "subscriptions(guid'#{@data['id']}')"
    end
  end

  class ChangeQuery < Sharepoint::Object
    include Sharepoint::Type
    belongs_to :list

    field 'Add',                   default: false
    field 'Alert',                 default: false
    field 'ChangeTokenEnd'
    field 'ChangeTokenStart'
    field 'ContentType',           default: false
    field 'DeleteObject',          default: false
    field 'Field',                 default: false
    field 'File',                  default: false
    field 'Folder',                default: false
    field 'Group',                 default: false
    field 'GroupMembershipAdd',    default: false
    field 'GroupMembershipDelete', default: false
    field 'Item',                  default: false
    field 'List',                  default: false
    field 'Move',                  default: false
    field 'Navigation',            default: false
    field 'Rename',                default: false
    field 'Restore',               default: false
    field 'RoleAssignmentAdd',     default: false
    field 'RoleAssignmentDelete',  default: false
    field 'RoleDefinitionAdd',     default: false
    field 'RoleDefinitionDelete',  default: false
    field 'RoleDefinitionUpdate',  default: false
    field 'SecurityPolicy',        default: false
    field 'Site',                  default: false
    field 'SystemUpdate',          default: false
    field 'Update',                default: false
    field 'User',                  default: false
    field 'View',                  default: false
    field 'Web',                   default: false

    def create
      @site.query :post, create_uri, { 'query' => @data }.to_json
    end
  end

  class ChangeItem < Sharepoint::Object
    field 'ChangeToken'
    field 'ChangeType'
    field 'ItemId'
    field 'ListId'
    field 'SiteId'
    field 'Time'
    field 'WebId'
  end

  class ChangeToken < Hash
    def initialize(string_value = nil, data = nil)
      super()
      self['__metadata'] = { 'type' => 'SP.ChangeToken' }
      self['StringValue'] = data.nil? ? string_value : data['StringValue']
    end
  end
end
