import discord
from discord.ext import commands
import openpyxl
from Apikeys import Discord  # Ensure your API key is imported correctly

# Define the command prefix
bot_prefix = '~'

# Define intents
intents = discord.Intents.default()
intents.messages = True
intents.guilds = True
intents.reactions = True
intents.message_content = True
intents.presences = True
intents.members = True

client = commands.Bot(command_prefix=bot_prefix, intents=intents)

# Role ID for the 'verified' role
verified_role_id = 1254790767327707136

# Load the Excel file
file_path = 'studycornerAllmember.xlsx'
wb = openpyxl.load_workbook(file_path)
ws = wb.active

# Function to find student name by student ID
def get_student_name(id):
    def get_row_number_by_value_in_column(workbook, sheet_name, column_letter, target_value):
        # Select the sheet by name
        sheet = workbook[sheet_name]

        # Get the column index from the column letter
        column_index = openpyxl.utils.column_index_from_string(column_letter)

        # Iterate through the rows in the specified column and find the target value
        for row_num, cell in enumerate(sheet.iter_rows(min_col=column_index, max_col=column_index)):
            if cell[0].value == target_value:
                # Return the row number (1-indexed)
                return row_num + 1

        # If the target value is not found, return None
        return None

    # Example usage
    sheet_name = 'Sheet1'  # Ensure the sheet name matches the actual sheet name in your file
    column_letter = 'A'
    target_value = int(id)

    row_number = get_row_number_by_value_in_column(wb, sheet_name, column_letter, target_value)

    if row_number is not None:
        sheet = wb[sheet_name]
        cell1 = f"B{row_number}"
        name = sheet[cell1].value
        return [True, name]
    else:
        return [False,None]

# Example call to the function with the specified ID
get_student_name(213588)

@client.event
async def on_ready():
    print("Bot is ready")
    print('---------------')

@client.command()
async def reg(ctx, student_id: str):
    student = get_student_name(student_id)
    student_name = student[1]
    valid_s = student[0]
    if valid_s:
        role = ctx.guild.get_role(verified_role_id)
        if role is not None:
            try:
                # Assign the 'verified' role to the member
                await ctx.author.add_roles(role)
                # Change the member's nickname to the corresponding name
                await ctx.author.edit(nick=student_name)
                embed = discord.Embed(title='Verified', url='https://www.facebook.com/BRACUCC',
                                      description='Welcome to BUCC Study Corner!', color=0x47d147)
                embed.set_author(name='BUCC', url='https://www.facebook.com/BRACUCC',
                                 icon_url='https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTh2qXgZYrSGqgQNigRaYXyYlPzwydhu7i-2g&s')
                embed.set_thumbnail(
                    url='https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTh2qXgZYrSGqgQNigRaYXyYlPzwydhu7i-2g&s')
                embed.add_field(name='ID', value=student_id, inline=True)
                embed.add_field(name='Name', value=student_name, inline=True)
                await ctx.send(embed=embed)
                # await ctx.send(f"You have been verified and your nickname has been changed to '{student_name}'.")
            except discord.Forbidden as e:
                print(f"Forbidden error: {e}")
                await ctx.send("I don't have permission to assign roles or change nicknames.")
            except discord.HTTPException as e:
                # Log detailed information about the HTTP exception
                print(f"HTTP exception details: Status: {e.status}, Code: {e.code}, Text: {e.text}")
                await ctx.send(f"An error occurred while processing your request. Please try again later.")
        else:
            await ctx.send("The 'verified' role does not exist.")
    else:
        await ctx.send(f"Student ID not found. Please provide a valid Student ID.\nPlease contact with your departmental seniors for further instruction")

# Error handling example
@client.event
async def on_command_error(ctx, error):
    if isinstance(error, commands.CommandNotFound):
        await ctx.send("Invalid command. Use ~reg <Student ID> to register.")
    elif isinstance(error, commands.BadArgument):
        await ctx.send("Please format your message correctly. Use ~reg <Student ID>.")
    elif isinstance(error, discord.Forbidden):
        print(f"Forbidden error: {error}")
        await ctx.send("I don't have the necessary permissions to perform this action.")
    else:
        print(f"Unhandled error: {error}")
        await ctx.send("An error occurred. Please try again.")
@client.command()
@commands.has_role('Panel Member')
async def announce(ctx, msg: str):
    channel = client.get_channel(765667602294767666)
    await channel.send(msg)

@client.command()
@commands.has_role('Panel Member')
async def ver(ctx, msg: str):
    channel = client.get_channel(1093479269683970118)
    await channel.send(msg)

@client.command()
@commands.has_role('Panel Member')
async def test(ctx, msg: str):
    channel = client.get_channel(1254810239417454622)
    await channel.send(msg)


client.run(Discord)
