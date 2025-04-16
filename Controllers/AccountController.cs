using ManagerPdf.Data;
using ManagerPdf.Models;
using ManagerPdf.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;

namespace ManagerPdf.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AccountController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private LdapService _ldapService;
        private readonly IConfiguration _configuration;

        public AccountController(ApplicationDbContext context, LdapService ldapService, IConfiguration configuration)
        {
            _context = context;
            _ldapService = ldapService;

            _configuration = configuration;
        }


        [HttpPost("login")]
        public async Task<IActionResult> Login()
        {
            string username = Request.Form["username"];
            string password = Request.Form["password"];

            string response = _ldapService.Login(username, password);

            var token = GenerateJwtToken(username);

            return Ok(new { message = response, token = token });
        }

        private string GenerateJwtToken(string username)
        {

            var jwtKey = _configuration["Jwt:Key"];
            var jwtIssuer = _configuration["Jwt:Issuer"];
            var jwtAudience = _configuration["Jwt:Audience"];

            var securityKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtKey));
            var credentials = new SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256);

            var claims = new[]
            {
                new Claim(JwtRegisteredClaimNames.Sub, username),
                new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
            };

            var token = new JwtSecurityToken(
                issuer: jwtIssuer,
                audience: jwtAudience,
                claims: claims,
                expires: DateTime.Now.AddHours(1),
                signingCredentials: credentials
            );

            return new JwtSecurityTokenHandler().WriteToken(token);
        }

        [HttpPost("getGroups")]
        public async Task<IActionResult> getGroups()
        {
            string username = Request.Form["username"];
            string password = Request.Form["password"];

            List<string> response = _ldapService.GetGroupNames(username, password);

            return Ok(new { message = response });
        }



        [HttpGet("getUsers")]
        public async Task<IActionResult> GetUsers()
        {
            List<Object> users = new List<Object>();
            var UsersDB = _context.Users.ToList();
            foreach (var User in UsersDB)
            {
                var userAdd = new
                {
                    User.Name, 
                    User.Email, 
                    User.Password,
                };

                users.Add(userAdd);
            }

            return Ok(users);
        }

        [HttpGet("getUsersDirectory")]
        public async Task<IActionResult> GetUsersDirectory()
        {
            string username = Request.Form["username"];
            string password = Request.Form["password"];

            string response = _ldapService.Login(username, password);

            return Ok(response);
        }

        [HttpPost("createUser")]
        public async Task<IActionResult> AddUsers()
        {
            string name = Request.Form["name"];
            string email = Request.Form["email"];
            string password = Request.Form["password"];

            var user = new ApplicationUser()
            {
                Name = name,
                Email = email,
                Password = password
            };
            await _context.Users.AddAsync(user);
            await _context.SaveChangesAsync();

            return Ok(new {message = "usuario regitrado correctamente"});
        }


        [HttpPost("updateUser/{id}")]
        public async Task<IActionResult> UpdateUsers(int id)
        {
            var userDB = await _context.Users.FirstOrDefaultAsync(x=>x.Id == id);
            

            string name = Request.Form["name"];
            string email = Request.Form["email"];
            string password = Request.Form["password"];

            userDB.Name = name;
            userDB.Email = email;
            userDB.Password = password;

            await _context.SaveChangesAsync();
            return Ok(new { message = "usuario actualizado correctamente" });
        }

    }
}
